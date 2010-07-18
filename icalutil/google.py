#!/usr/bin/env python


import time
import os.path
import errno

import vobject
import gdata.calendar
import gdata.calendar.service
import atom


def getdtstr(vevent, attrname):
    '''Return localized formatted time string for DTSTART or DTEND values.'''
    vdt = getattr(vevent, attrname)
    if hasattr(vdt.value, 'time'):
        dt = vdt.value   # datetime
        if not hasattr(dt, 'tzinfo'):
            raise AttributeError('No timezone information', vevent)
        return time.strftime('%Y-%m-%dT%H:%M:%S.000Z', dt.utctimetuple())
    # Date only ("all-day" event)
    return time.strftime('%Y-%m-%d', vdt.value.timetuple())


def createCalendarEventEntry(vevent):
    if hasattr(vevent, 'rrule'):
        children = [
            vevent.dtstart,
            vevent.dtend,
            vevent.rrule,
            ]
        if hasattr(vevent, 'exdate'):
            children.extend([child for child in vevent.getChildren()
                if child.name == 'EXDATE'])
        recurrence = gdata.calendar.Recurrence(
            # 'DTSTART;TZID=America/Los_Angeles:20050603T080000\r\n'
            # 'DTEND;TZID=America/Los_Angeles:20050603T080000\r\n'
            # 'RRULE:FREQ=DAILY;INTERVAL=1;UNTIL=20050609T065959Z;WKST=SU\r\n'
            # 'EXDATE;TZID=America/Los_Angeles:20091204T070000\r\n'
            # 'EXDATE;TZID=America/Los_Angeles:20091205T070000\r\n'
            text = ''.join([child.serialize() for child in children])
            )
        when = []
    else:
        recurrence = None
        when = [
            gdata.calendar.When(
                start_time = getdtstr(vevent, 'dtstart'),
                end_time = getdtstr(vevent, 'dtend'),
                ),
            ]

    transparencyValue = vevent.getChildValue('transp')
    if transparencyValue:
        transparency = gdata.calendar.Transparency()
        transparency.value = transparencyValue
    else:
        transparency.value = None

    uidValue = vevent.getChildValue('uid')
    if uidValue:
        uid = gdata.calendar.UID(
            value = uidValue,
            )
    else:
        uid = None
    return gdata.calendar.CalendarEventEntry(
        title = atom.Title(
            text = vevent.getChildValue('summary', '').strip()
            ),
        content = atom.Content(
            text = vevent.getChildValue('description', '').strip()
            ),
        where = [
            gdata.calendar.Where(
                value_string = vevent.getChildValue('location', '').strip()
                ),
            ],
        recurrence = recurrence,
        when = when,
        transparency = transparency,
        uid = uid,
        )


class uploader:

    def __init__(self,
            username = None,
            password = None,
            source = 'icalutil',
            calendar_id = None,
            dry_run = False,
            fail_dir = None,
            ):
        for dirname in [fail_dir]:
            if dirname and not os.path.isdir(dirname):
                raise EnvironmentError(errno.ENOENT, os.strerror(errno.ENOENT),
                    dirname)
        if not username:
            raise EnvironmentError('username is required')
        if not password:
            raise EnvironmentError('password is required')
        if not calendar_id:
            calendar_id = 'default'
        self.username = username
        self.password = password
        self.source = source
        self.cal = None
        self.upload_uri = '/calendar/feeds/%s/private/full' % calendar_id
        self.fail_dir = fail_dir
        self.dry_run = dry_run
            
    def uploadcalendar(self, ical,
            filteropts = None,
            eventcallbacks = None,
            ):
        if eventcallbacks is None:
            eventcallbacks = {}
        failed = []
        for component in ical.components():
            if component.name == vobject.icalendar.VEvent.name:
                try:
                    if not self.uploadevent(
                            vevent = component,
                            filteropts = filteropts,
                            eventcallbacks = eventcallbacks,
                            ):
                        failed.append(component)
                except gdata.service.RequestError, e:
                    if eventcallbacks.get('eventfailed'):
                        eventcallbacks.get('eventfailed')(self, component,
                            eventcallbacks.get('eventfailedarg'), e)
                        failed.append(component)
                        continue
                    raise
        return failed

    def uploadevent(self, vevent,
            filteropts = None,
            eventcallbacks = None,
            ):
        if eventcallbacks is None:
            eventcallbacks = {}
        entry = createCalendarEventEntry(vevent)
        if filteropts and filteropts.get('filter') and \
                not filteropts.get('filter')(vevent, entry,
                    filteropts.get('opts')):
            return False
        try:
            while True:
                try:
                    if not self.cal:
                        if eventcallbacks.get('beforelogin'):
                            eventcallbacks.get('beforelogin')()
                        cal = gdata.calendar.service.CalendarService()
                        if not self.dry_run:
                            cal.ClientLogin(
                                username = self.username,
                                password = self.password,
                                source = self.source,
                                )
                        self.cal = cal
                    if eventcallbacks.get('beforeinsert'):
                        eventcallbacks.get('beforeinsert')(self, vevent, entry,
                            eventcallbacks.get('beforeinsertarg'))
                    if not self.dry_run:
                        self.cal.InsertEvent(entry, self.upload_uri)
                    break
                except gdata.service.RequestError, e:
                    if eventcallbacks.get('eventexception'):
                        eventcallbacks.get('eventexception')(self, vevent,
                            entry, e)
                        continue
                    raise
        finally:
            if eventcallbacks.get('afterinsert'):
                eventcallbacks.get('afterinsert')(self, vevent, entry,
                    eventcallbacks.get('afterinsertarg'))
        return True
