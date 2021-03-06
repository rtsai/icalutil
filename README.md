This project contains the following modules:

icalutil
========

This is a layer on top of [vobject] and provides some utility functions for
traversing and manipulating iCalendar objects and files.

  [vobject]: http://vobject.skyhouseconsulting.com/

icalutil.google
===============

This provides some functions for mapping vobjects to Google Calendar event
entry objects and calling the Google Calendar API.

Features
--------

- All events uploaded as UTC; timezone information should be preserved.
- Supports recurring events (`RRULE`) with exceptions.

Known Bugs
----------

- `RDATE` and `EXRULE` components are not fully supported.
- Google Calendar only provides about 720 recurrences for recurring events (or
  two years, for a DAILY event). A workaround would be to duplicate the
  recurring event until the present time is reached.
- Events with many exceptions (empirically observed to be around 72 or more
  `EXDATE` children) are rejected by the Google Calendar API with an error
  message of "`RDATE too large`"; the `icalutil` tools do not provide any
  workaround.
- The Google Calendar API [batch requests] mechanism is not used.

  [batch requests]: http://code.google.com/apis/calendar/data/2.0/developers_guide_protocol.html#batch

gcalfiltersplit
===============

Split up a single ICS file into smaller ones more manageable by the Google
Calendar 'import' tool.

The motivation for writing this was to use the Google Calendar 'import' tool to
import smaller .ics files. However, it turns out that this 'import' is also
subject to Google Calendar API quotas, so `gcalfiltersplit` is not recommended
for bulk upload, since the error reporting is not very good.

For a dry run:

    ./gcalfiltersplit -n ical.ics

For the real thing:

    ./gcalfiltersplit ical.ics

gcaluploader
============

Read an iCalendar file and upload its events into Google Calendar, with
comprehensive logging and error reporting.

- Tracks and logs failed uploads so that individual entries may be examined
  offline for re-upload or manual entry.
- Sleep/retry recovery from transient errors:
    - HTTP 302 redirects.
    - Google Calendar API call quotas:
	- "Burst" rate appears to be approximately 4000 API calls in one day.
	- Sustained rate appears to be an average of 1 API call every 10
	  seconds.
- Optional sanitization of calendar entries:
    - Coalesce recurring daily all-day events into a single multi-day event.
    - Filter events with empty summary strings.
    - Workarounds for buggy Apple iCal.app import of Palm Desktop vCal export.

For a dry run:

    ./gcaluploader -n ical.ics

For the real thing:

    ./gcaluploader ical.ics


Credits
=======

Inspiration from [ics-gcal.py]

  [ics-gcal.py]: http://repo.ub3rgeek.net/branches/misc-scripts/annotate/head:/ics-gcal.py
