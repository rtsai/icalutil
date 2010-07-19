#!/usr/bin/env python


import optparse
import ConfigParser
import os.path
import time
import datetime
import re
import pytz
import calendar
import copy
import errno

import vobject
import gdata.calendar
import icalutil
import icalutil.google


def getconfigstr(config, fieldname):
    try:
        return config.get(ConfigParser.DEFAULTSECT, fieldname)
    except ConfigParser.NoOptionError:
        pass


def getconfigboolean(config, fieldname):
    try:
        return config.getboolean(ConfigParser.DEFAULTSECT, fieldname)
    except ConfigParser.NoOptionError:
        pass


def getconfigint(config, fieldname):
    try:
        return config.getint(ConfigParser.DEFAULTSECT, fieldname)
    except ConfigParser.NoOptionError:
        pass


def getboolopt(options, config, fieldname):
    val = getattr(options, fieldname)
    if val is not None:
        return val
    return getconfigboolean(config, fieldname)


def filterevent(vobj, opts):
    '''Filter for iCalendar VEVENT components.'''
    if vobj.name == vobject.icalendar.VCalendar2_0.name:
        # VCALENDAR has VEVENT children
        return True
    if vobj.name == vobject.icalendar.VEvent.name:
        memo = opts['memo']
        filters = memo['filters']
        transforms = memo['transforms']
        uid = vobj.getChildValue('uid')
        start_uid = opts.get('start_uid')
        if start_uid:
            if uid != start_uid:
                return False
            del opts['start_uid']
        if not opts['preserve_uids'] and hasattr(vobj, 'uid'):
            del vobj.uid
        if opts['select_uids'] and uid not in opts['select_uids']:
            # Select UIDs
            return False
        if not opts['accept_empty_summary'] and \
                not vobj.getChildValue('summary', '').strip():
            # Reject events with empty summary strings.
            filters[uid] = 'empty summary'
            return False
        rrule = vobj.getChildValue('rrule')
        if rrule:
            # Reject events that run forever (buggy Apple iCal Palm vCal
            # import).
            rruleparams = dict([kvp.upper().split('=', 1)
                for kvp in rrule.split(';')])
            freq = rruleparams[u'FREQ']
            if u'UNTIL' not in rruleparams and \
                    freq not in opts['accept_neverending_recurrences']:
                filters[uid] = 'unending %s recurrence' % freq
                return False
        if opts['enable_vcal_import_workaround_hack']:
            # All-day events (recurring and non-recurring) in a Palm vCal
            # export get imported by Apple iCal as a two-day event starting one
            # day early. Undo that miscalculation.
            dtstart = vobj.getChildValue('dtstart')
            dtend = vobj.getChildValue('dtend')
            if not hasattr(dtstart, 'time') and not hasattr(dtend, 'time'):
                td = dtend - dtstart
                if td.days == 2 and td.seconds == 0 and td.microseconds == 0:
                    vobj.dtstart.value = dtstart + datetime.timedelta(days = 1)
                    if not transforms.get(uid):
                        transforms[uid] = []
                    transforms[uid].append('vcal-import-workaround')
        if opts['coalesce_events']:
            if icalutil.coalesce(vobj, True):
                if not transforms.get(uid):
                    transforms[uid] = []
                transforms[uid].append('coalesced %d days' %
                    (vobj.getChildValue('dtend') -
                    vobj.getChildValue('dtstart')).days)
        if opts['truncate_exdates']:
            # Keep the newest exdate children
            exdates = [child for child in vobj.getChildren()
                if child.name == u'EXDATE']
            deco = [(exdate.value[0], exdate) for exdate in exdates]
            deco.sort()
            deco.reverse()      # descending date
            remove = [d[1] for d in deco[opts['truncate_exdates']:]]
            if remove:
                for child in remove:
                    vobj.remove(child)
                if not transforms.get(uid):
                    transforms[uid] = []
                transforms[uid].append('truncated oldest %d exdate(s)' %
                    len(remove))
        if opts['max_exdates']:
            exdates = [child for child in vobj.getChildren()
                if child.name == u'EXDATE']
            if len(exdates) > opts['max_exdates']:
                filters[uid] = '%d EXDATEs, max=%d' % \
                    (len(exdates), opts['max_exdates'])
                return False
        return True
    # Discard everything else
    return False


def noop(*args, **kwargs):
    pass


def log(msg):
    print '%s: %s' % (datetime.datetime.now().strftime('%H:%M:%S'), msg)


def beforelogin():
    log('Logging in ...')


NEWLINE_RE = re.compile('[\r\n]')

def beforeinsert(uploader, vevent, entry, uploadmemo):
    __pychecker__ = 'unusednames=uploader,vevent,entry'
    split = NEWLINE_RE.split(entry.title.text, 1)
    if split:
        title = split[0].strip()
    else:
        title = None
    uid = vevent.getChildValue('uid') or 'No UID'
    msg = 'Inserting %d/%d: %s (%s)' % (uploadmemo['inserts'] + 1,
        uploadmemo['end'], uid, title)
    if entry.when:
        msg += ' (%s)' % entry.when[0].start_time
    elif entry.recurrence:
        msg += ' (%s)' % \
            ';'.join(
            ';'.join(entry.recurrence.text.split('\r\n')[
            0:3:2]).replace('TZID=America/Los_Angeles:', '').
            split(';')[0:3])
    if uid:
        reasons = uploadmemo['transforms'].get(uid)
        if reasons:
            msg += ' (%s)' % ','.join(reasons)
    log(msg)


def afterinsert(uploader, vevent, entry, uploadmemo):
    __pychecker__ = 'unusednames=uploader,vevent,entry'
    uploadmemo['inserts'] += 1


def eventexception(uploader, vevent, entry, e):
    __pychecker__ = 'unusednames=vevent,entry'
    eargs = e.args[0]
    if eargs['status'] == 403 and \
            eargs['reason'] == 'Forbidden' and \
            eargs['body'] == 'The user has exceeded their quota, and cannot ' \
                'currently perform this operation':
        uploader.cal = None
        log(e)
        log('Sleeping for 5 minutes')
        time.sleep(5 * 60)
        return
    if eargs['status'] == 302:
        log(e)
        log('Sleeping for 5 seconds')
        time.sleep(5)
        return
    if eargs['status'] == 500 and \
            eargs['reason'] == 'Internal Server Error' and \
            eargs['body'] == 'Service error: could not insert entry':
        log(e)
        log('Sleeping for 5 seconds')
        time.sleep(5)
        return
    raise e


def eventfailed(uploader, vevent, uploadmemo, e):
    __pychecker__ = 'unusednames=uploader'
    uid = vevent.getChildValue('uid')
    eargs = e.args[0]
    if eargs['reason'] == 'Conflict':
        msg = eargs['reason']
    else:
        msg = str(e)
    uploadmemo['fails'][uid] = msg
    log('Failed UID: %s (%s)' % (uid, msg))
    if eargs['status'] == 400 or \
            eargs['status'] == 409 and eargs['reason'] == 'Conflict':
        if not uploader.dry_run and uploader.fail_dir:
            filename = os.path.join(uploader.fail_dir, uid + '.ics')
            f = open(filename, 'w')
            try:
                newical = vobject.iCalendar()
                newical.add(vevent)
                f.write(newical.serialize())
            finally:
                f.close()
        return
    raise e


def filterentry(vevent, entry, opts):
    '''Filter entries before uploading to Google.'''
    if opts and opts['reminder_minutes'] is not None and \
            ('valarm' in vevent.contents or opts['force_reminder']):
        reminder = gdata.calendar.Reminder(
            minutes = opts['reminder_minutes'],
            )
        for when in entry.when:
            when.reminder.append(reminder)
    return True


def gettz(components):
    timezones = [c for c in components
        if c.name == vobject.icalendar.VTimezone.name]
    if timezones:
        if len(timezones) > 1:
            raise AttributeError('Multi-timezone calendars are not supported',
                timezones)
        tzid = timezones[0].getChildValue('tzid')
        if tzid:
            return pytz.timezone(tzid)
        return pytz.utc
    return pytz.utc


def componentdt(component, tz):
    '''Return a timestamp for sorting.'''
    d = component.getChildValue('dtstart')
    if d is None:
        tm = datetime.datetime.utcfromtimestamp(0).utctimetuple()
    else:
        if hasattr(d, 'time'):
            if not hasattr(d, 'tzinfo'):
                raise AttributeError('No timezone information', component)
            tm = d.utctimetuple()
        else:
            tm = tz.localize(datetime.datetime.combine(d,
                datetime.time())).utctimetuple()
    return calendar.timegm(tm)


def reportuids(vevents, uids, reasons, verb):
    if uids:
        log('%s %d UIDs (selecting %d UIDs)' % (verb, len(vevents), len(uids)))
    for uid in [vobj.getChildValue('uid') for vobj in vevents]:
        if uid in reasons:
            log('%s UID: %s (%s)' % (verb, uid, reasons[uid]))


def optparse_setdefaults(p, config):
    '''
    Set optparse defaults from ConfigParser defaults.

    Desired behavior: optparse defaults, ConfigParser, optparse command-line;
    last one wins.

    The problem is that optparse defaults are applied after command-line
    parsing, and there is no way to know how an option got set (command-line,
    optparse defaults) before applying ConfigParser defaults.

    So the actual behavior is:

    ConfigParser, optparse command-line; last one wins. There are no optparse
    defaults, so the help output is suboptimally helpful.
    '''
    return
    options = dict(zip(
        [option.dest for option in p.option_list],
        p.option_list,
        ))
    for key, value in config.defaults().iteritems():
        option = options.get(key)
        if option:
            if option.action == 'store':
                if option.type == 'string':
                    p.set_default(key, value)
                elif option.type == 'int':
                    p.set_default(key, int(value))
                elif option.type == 'long':
                    p.set_default(key, long(value))
                elif option.type == 'float':
                    p.set_default(key, float(value))
                elif option.type == 'complex':
                    p.set_default(key, complex(value))
                else:
                    raise Exception('Unknown optparse type', key, value,
                        option.type)
            elif option.action in ['store_true', 'store_false']:
                p.set_default(key, config._boolean_states[value.lower()])
            else:
                raise Exception('Unknown optparse action', key, value,
                    option.action)


def getoptions(description, config_file, config_vars):
    p = optparse.OptionParser(
        description = description,
        )
    config = ConfigParser.SafeConfigParser()
    if 'config_file' in config_vars:
        p.add_option('-c', '--config-file',
            dest = 'config_file',
            metavar = 'FILENAME',
            help = 'Configuration file for options (default: %default)',
            )
    if 'username' in config_vars:
        p.add_option('-u', '--username',
            dest = 'username',
            metavar = 'EMAIL_ADDRESS',
            help = 'Gmail or Google Apps e-mail address (default: %default)',
            )
    if 'password' in config_vars:
        p.add_option('-p', '--password',
            dest = 'password',
            help = 'Account password (default: %default)',
            )
    if 'calendar_id' in config_vars:
        p.add_option('-i', '--calendar-id',
            dest = 'calendar_id',
            help = 'Google Calendar ID (default: %default)',
            )
        config.set(ConfigParser.DEFAULTSECT, 'calendar_id', 'default')
    if 'quiet' in config_vars:
        p.add_option('-q', '--quiet',
            dest = 'quiet',
            action = 'store_true',
            help = 'Suppress output (default: %default)',
            )
        config.set(ConfigParser.DEFAULTSECT, 'quiet', 'false')
    if 'dry_run' in config_vars:
        p.add_option('-n', '--dry-run',
            dest = 'dry_run',
            action = 'store_true',
            help = 'Don\'t actually upload anything (default: %default)',
            )
        config.set(ConfigParser.DEFAULTSECT, 'dry_run', 'false')
    if 'max_filesize' in config_vars:
        p.add_option('-m', '--max-filesize',
            type = 'int',
            dest = 'max_filesize',
            help = 'Split input file into files of approximate maximum size',
            )
        config.set(ConfigParser.DEFAULTSECT, 'max_filesize', '524288')
    if 'fail_dir' in config_vars:
        p.add_option('--fail-dir',
            dest = 'fail_dir',
            metavar = 'DIRECTORY',
            help = 'Directory to receive not-uploaded .ics files ' \
                '(default: %default)',
            )
    if 'reminder_minutes' in config_vars:
        p.add_option('-r', '--reminder-minutes',
            dest = 'reminder_minutes',
            type = 'int',
            metavar = 'MINUTES',
            help = 'Reminder time for event (default: %default)',
            )
        config.set(ConfigParser.DEFAULTSECT, 'reminder_minutes', '30')
    if 'force_reminder' in config_vars:
        p.add_option('-R', '--force-reminder',
            dest = 'force_reminder',
            action = 'store_true',
            help = 'Force a reminder to be set for all events ' \
                '(default: %default)',
            )
        config.set(ConfigParser.DEFAULTSECT, 'force_reminder', 'false')
    if 'enable_vcal_import_workaround_hack' in config_vars:
        p.add_option('-V', '--disable-vcal-import-workaround-hack',
            dest = 'enable_vcal_import_workaround_hack',
            action = 'store_false',
            help = 'Disable iCal Palm vCal import hack (default %default)',
            )
        config.set(ConfigParser.DEFAULTSECT,
            'enable_vcal_import_workaround_hack', 'true')
    if 'start_uid' in config_vars:
        p.add_option('--start-uid',
            dest = 'start_uid',
            help = 'Only start uploading events starting with UID',
            )
        config.set(ConfigParser.DEFAULTSECT, 'start_uid', '')
    if 'select_uids' in config_vars:
        p.add_option('--select-uids',
            dest = 'select_uids',
            help = 'Only select events with the given UIDs (comma-delimited)',
            )
        config.set(ConfigParser.DEFAULTSECT, 'select_uids', '')
    if 'preserve_uids' in config_vars:
        p.add_option('-I', '--disable-preserve-uids',
            dest = 'preserve_uids',
            action = 'store_false',
            help = 'Ignore UID values in calendar file',
            )
        config.set(ConfigParser.DEFAULTSECT, 'preserve_uids', 'true')
    if 'coalesce_events' in config_vars:
        p.add_option('-C', '--disable-coalesce-events',
            dest = 'coalesce_events',
            action = 'store_false',
            help = 'Don\'t coalesce recurring daily events into a single ' \
                'multi-day event.',
            )
        config.set(ConfigParser.DEFAULTSECT, 'coalesce_events', 'true')
    if 'truncate_exdates' in config_vars:
        p.add_option('--truncate-exdates',
            type = 'int',
            dest = 'truncate_exdates',
            help = 'Keep the newest TRUNCATE-EXDATE recurrence exceptions, ' \
                'discarding the older ones.',
            )
        config.set(ConfigParser.DEFAULTSECT, 'truncate_exdates', '0')
    if 'max_exdates' in config_vars:
        p.add_option('--max-exdates',
            type = 'int',
            dest = 'max_exdates',
            help = 'Accept events with up to MAX-EXDATE recurrence exceptions.',
            )
        config.set(ConfigParser.DEFAULTSECT, 'max_exdates', '72')
    if 'accept_neverending_recurrences' in config_vars:
        p.add_option('-N', '--accept-neverending-recurrences',
            dest = 'accept_neverending_recurrences',
            help = 'Accept certain kinds of neverending recurrences ' \
                '(default: %default)',
            )
        config.set(ConfigParser.DEFAULTSECT, 'accept_neverending_recurrences',
            'daily,weekly,monthly,yearly')
    if 'accept_empty_summary' in config_vars:
        p.add_option('-S', '--accept-empty-summary',
            action = 'store_true',
            dest = 'accept_empty_summary',
            help = 'Accept events with empty summaries (default: %default)',
            )
        config.set(ConfigParser.DEFAULTSECT, 'accept_empty_summary', 'false')
    p.set_defaults(
        config_file = config_file,
        )
    optparse_setdefaults(p, config)
    options, args = p.parse_args()

    config.read([
        options.config_file,
        os.path.expanduser('~/.' + config_file),
        ])

    opts = {}
    if 'username' in config_vars:
        opts['username'] = options.username or getconfigstr(config, 'username')
    if 'password' in config_vars:
        opts['password'] = options.password or getconfigstr(config, 'password')
    if 'calendar_id' in config_vars:
        opts['calendar_id'] = options.calendar_id or \
            getconfigstr(config, 'calendar_id')
    if 'quiet' in config_vars:
        opts['quiet'] = getboolopt(options, config, 'quiet')
    if 'dry_run' in config_vars:
        opts['dry_run'] = getboolopt(options, config, 'dry_run')
    if 'max_filesize' in config_vars:
        opts['max_filesize'] = options.max_filesize or \
            getconfigint(config, 'max_filesize')
    if 'fail_dir' in config_vars:
        opts['fail_dir'] = options.fail_dir or getconfigstr(config, 'fail_dir')
    if 'reminder_minutes' in config_vars:
        opts['reminder_minutes'] = options.reminder_minutes or \
            getconfigint(config, 'reminder_minutes')
    if 'force_reminder' in config_vars:
        opts['force_reminder'] = getboolopt(options, config, 'force_reminder')
    if 'enable_vcal_import_workaround_hack' in config_vars:
        opts['enable_vcal_import_workaround_hack'] = getboolopt(options, config,
            'enable_vcal_import_workaround_hack')
    if 'start_uid' in config_vars:
        opts['start_uid'] = (options.start_uid or
            getconfigstr(config, 'start_uid') or '').strip().upper()
    if 'select_uids' in config_vars:
        opts['select_uids'] = dict([(x.strip().upper(), True)
            for x in (options.select_uids or getconfigstr(config, 'select_uids')
                    or '').split(',') if x])
    if 'preserve_uids' in config_vars:
        opts['preserve_uids'] = getboolopt(options, config, 'preserve_uids')
    if 'coalesce_events' in config_vars:
        opts['coalesce_events'] = getboolopt(options, config, 'coalesce_events')
    if 'truncate_exdates' in config_vars:
        opts['truncate_exdates'] = options.truncate_exdates or \
            getconfigint(config, 'truncate_exdates')
    if 'max_exdates' in config_vars:
        opts['max_exdates'] = options.max_exdates or \
            getconfigint(config, 'max_exdates')
    if 'accept_neverending_recurrences' in config_vars:
        opts['accept_neverending_recurrences'] = [x.strip().upper()
            for x in (options.accept_neverending_recurrences or
            getconfigstr(config, 'accept_neverending_recurrences')).split(',')]
    if 'accept_empty_summary' in config_vars:
        opts['accept_empty_summary'] = getboolopt(options, config,
            'accept_empty_summary')
    return opts, args


def splitcb(cal, arg):
    dirname, basename = os.path.split(arg['filename'])
    basenameprefix, basenameext = os.path.splitext(basename)
    newpath = os.path.join(
        dirname,
        basenameprefix + ('-%06d' % arg['splits']) + basenameext,
        )
    arg['splits'] += 1
    if os.path.exists(newpath):
        raise EnvironmentError(errno.EEXIST, os.strerror(errno.EEXIST), newpath)
    events = [c for c in cal.components()
        if c.name == vobject.icalendar.VEvent.name]
    splitmemo = arg['splitmemo']
    for event in events:
        uid = event.getChildValue('uid')
        reasons = splitmemo['transforms'].get(uid)
        if uid and reasons:
            log('Transformed UID %s: %s' % (uid, ','.join(reasons)))
    newsize = 0
    if not arg['dry_run']:
        f = open(newpath, 'w')
        try:
            f.write(cal.serialize())
        finally:
            f.close()
        newsize = os.path.getsize(newpath)
    log('Wrote %s: events=%d, bytes=%d' % (newpath, len(events), newsize))


def filtersplit():
    opts, args = getoptions(
        description = 'Split a single iCal file into multiple smaller files',
        config_file = 'gcalfiltersplit.cnf',
        config_vars = [
            'config_file',
            'quiet',
            'dry_run',
            'max_filesize',

            'enable_vcal_import_workaround_hack',
            'start_uid',
            'select_uids',
            'preserve_uids',
            'coalesce_events',
            'truncate_exdates',
            'max_exdates',
            'accept_neverending_recurrences',
            'accept_empty_summary',
            ],
        )
    if not args:
        print 'No files!'
        return 1

    if opts['quiet']:
        global log
        log = noop

    for filename in args:
        f = open(filename)
        try:
            log('Reading %s ...' % filename)
            ical = vobject.readComponents(f).next()
            components = [c for c in ical.components()]
            nevents = len([c for c in components
                if c.name == vobject.icalendar.VEvent.name])
            log('Sorting %d events by descending date ...' % nevents)
            tz = gettz(components)
            deco = [(componentdt(c, tz), c) for c in components]
            deco.sort()
            components = [pair[1] for pair in deco]
            components.reverse()    # Descending dtstart
            ical = icalutil.createcalendar(components)
            components = None
            deco = None
        finally:
            f.close()
        splitmemo = {
            'filters': {},
            'transforms': {},
        }
        filteropts = copy.copy(opts)
        filteropts['memo'] = splitmemo
        filtered = icalutil.filtercomponents(ical, filterevent, filteropts)
        reportuids(filtered, opts['select_uids'], splitmemo['filters'],
            'Filtered')
        newnevents = len([c for c in ical.components()
            if c.name == vobject.icalendar.VEvent.name])
        if not newnevents:
            log('No events!')
            return 0
        # Reduce filesize by filtered events
        oldfilesize = os.path.getsize(filename) * newnevents / nevents
        # Guess number of events in each sub-calendarfile
        splitevents = nevents * opts['max_filesize'] / oldfilesize
        start = int(time.time())
        arg = {
            'filename': filename,
            'splits': 0,
            'dry_run': opts['dry_run'],
            'splitmemo': splitmemo,
        }
        try:
            icalutil.splitcal(ical,
                events_per_calendar = splitevents,
                splitcallback = splitcb,
                splitcallbackarg = arg,
                )
        finally:
            log('Elapsed time: %d second(s)' % (int(time.time()) - start))

    return 0


def upload():
    opts, args = getoptions(
        description = 'Upload iCal .ics files to Google Calendar',
        config_file = 'gcaluploader.cnf',
        config_vars = [
            'config_file',
            'username',
            'password',
            'calendar_id',
            'quiet',
            'dry_run',
            'fail_dir',
            'reminder_minutes',
            'force_reminder',

            'enable_vcal_import_workaround_hack',
            'start_uid',
            'select_uids',
            'preserve_uids',
            'coalesce_events',
            'truncate_exdates',
            'max_exdates',
            'accept_neverending_recurrences',
            'accept_empty_summary',
            ],
        )
    if not args:
        print 'No files!'
        return 1

    if opts['quiet']:
        global log
        log = noop

    uploader = icalutil.google.uploader(
        username = opts['username'],
        password = opts['password'],
        calendar_id = opts['calendar_id'],
        dry_run = opts['dry_run'],
        fail_dir = opts['fail_dir'],
        )

    eventcallbacks = {}
    eventcallbacks['beforelogin'] = beforelogin
    eventcallbacks['beforeinsert'] = beforeinsert
    eventcallbacks['afterinsert'] = afterinsert
    eventcallbacks['eventexception'] = eventexception
    eventcallbacks['eventfailed'] = eventfailed

    for filename in args:
        f = open(filename)
        try:
            log('Reading %s ...' % filename)
            ical = vobject.readComponents(f).next()
            components = [c for c in ical.components()]
            nevents = len([c for c in components
                if c.name == vobject.icalendar.VEvent.name])
            log('Sorting %d events by descending date ...' % nevents)
            tz = gettz(components)
            deco = [(componentdt(c, tz), c) for c in components]
            deco.sort()
            components = [pair[1] for pair in deco]
            components.reverse()    # Descending dtstart
            ical = icalutil.createcalendar(components)
            components = None
            deco = None
        finally:
            f.close()
        uploadmemo = {
            'inserts': 0,
            'end': nevents,
            'fails': {},
            'filters': {},
            'transforms': {},
        }
        eventcallbacks['beforeinsertarg'] = uploadmemo
        eventcallbacks['afterinsertarg'] = uploadmemo
        eventcallbacks['eventfailedarg'] = uploadmemo
        filteropts = copy.copy(opts)
        filteropts['memo'] = uploadmemo
        filtered = icalutil.filtercomponents(ical, filterevent, filteropts)
        reportuids(filtered, opts['select_uids'], uploadmemo['filters'],
            'Filtered')
        nevents = len([c for c in ical.components()
            if c.name == vobject.icalendar.VEvent.name])
        uploadmemo['end'] = nevents
        failed = []
        start = int(time.time())
        try:
            failed = uploader.uploadcalendar(
                ical = ical,
                filteropts = {
                    'filter': filterentry,
                    'opts': {
                        'reminder_minutes': opts['reminder_minutes'],
                        'force_reminder': opts['force_reminder'],
                    },
                },
                eventcallbacks = eventcallbacks,
                )
        except gdata.service.CaptchaRequired:
            domain = opts['username'].split('@', 1)[1]
            if domain == 'gmail.com':
                path = 'accounts'
            else:
                path = 'a/%s' % domain    # Google Apps
            log('https://www.google.com/%s/UnlockCaptcha' % path)
            raise
        finally:
            reportuids(filtered, opts['select_uids'], uploadmemo['filters'],
                'Filtered')
            reportuids(filtered, None, uploadmemo['transforms'], 'Transformed')
            for uid in [vevent.getChildValue('uid') for vevent in failed]:
                log('Failed UID: %s (%s)' % (uid, uploadmemo['fails'][uid]))
            log('Inserted %d event(s)' % uploadmemo['inserts'])
            log('Elapsed time: %d second(s)' % (int(time.time()) - start))

    return 0
