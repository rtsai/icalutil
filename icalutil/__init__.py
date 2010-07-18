#!/usr/bin/env python


import collections
import vobject
import datetime
import copy


def walkcomponents(vobj, f, arg):
    '''Walk a component tree (breadth-first).'''
    components = collections.deque([vobj,])
    while components:
        component = components.popleft()
        f(component, arg)
        components.extend([component for component in vobj.components()])


def filtercomponents(vobj, f, arg):
    '''
    Filter a component tree (breadth-first); remove components when 'f' returns
    False. Return the list of removed components.
    '''
    filtered = []
    if not f(vobj, arg):
        raise Exception('Can\'t filter root component')
    components = collections.deque([vobj,])
    while components:
        component = components.popleft()
        remove = []
        for c in component.components():
            if f(c, arg):
                components.append(c)
            else:
                remove.append(c)
        for c in remove:
            component.remove(c)
        filtered.extend(remove)
    return filtered


def coalesce(vevent, inplace):
    '''
    Coalesce a recurring daily all-day event into a single non-recurring
    multi-day event.
    '''
    if vevent.name != vobject.icalendar.VEvent.name:
        return vevent                   # not a VEVENT component

    for dtattr in ['dtstart', 'dtend']:
        dtval = vevent.getChildValue(dtattr)
        if not dtval:
            return vevent
        if hasattr(dtval, 'time'):
            return vevent               # timestamped; not all-day

    if vevent.getChildValue('exdate'):
        return vevent                   # exceptions

    rruleValue = vevent.getChildValue('rrule')
    if not rruleValue:
        return vevent                   # no recurrence

    rruleparams = dict([kvp.upper().split('=', 1)
        for kvp in rruleValue.split(';')])
    if rruleparams.get(u'FREQ') != u'DAILY':
        return vevent
    if rruleparams.get(u'INTERVAL') != u'1':
        return vevent
    until = rruleparams.get(u'UNTIL')
    if not until:
        return vevent

    if inplace:
        newv = vevent
    else:
        newv = copy.deepcopy(vevent)

    newv.dtend.value = datetime.datetime.strptime(rruleparams.get(u'UNTIL'),
        '%Y%m%d').date()
    del newv.rrule
    return newv


def createcalendar(components):
    cal = vobject.iCalendar()
    for c in components:
        cal.add(c)
    return cal


def splitcal(cal,
        events_per_calendar = 1,
        splitcallback = None,
        splitcallbackarg = None,
        ):
    '''
    Read an iCalendar file and write its component events as separate calendar
    files in the given directory.
    '''
    nonevents = [c for c in cal.components()
        if c.name != vobject.icalendar.VEvent.name]
    newcal = createcalendar(nonevents)
    newcalevents = 0
    written = 0
    for vobj in cal.components():
        if vobj.name != vobject.icalendar.VEvent.name:
            continue
        newcal.add(vobj)
        newcalevents += 1
        if events_per_calendar > 0 and newcalevents >= events_per_calendar:
            if splitcallback:
                splitcallback(newcal, splitcallbackarg)
            written += 1
            newcal = createcalendar(nonevents)
            newcalevents = 0
    if newcalevents > 0:
        if splitcallback:
            splitcallback(newcal, splitcallbackarg)
