# GAS-Facebook-Event-Sync

This is a standalone script (that consists of multiple files). The purpose is to sync Facebook event calendars to Google Calendar.

This project is a fork of [GAS-ICS-Sync](https://github.com/derekantrican/GAS-ICS-Sync), specialized for Facebook's ICS exports. To obtain Facebook's ICS export, go to your [Facebook events calendar](https://www.facebook.com/events/calendar), right click on the "Add to calendar" icon, and click copy link address. This can be used as the calendar source.

For setup instructions, see the following guide at the upstream repo: https://github.com/derekantrican/GAS-ICS-Sync/wiki/Setting-up-the-script-manually.

## Excluding Canceled Facebook events

Facebook does not mark canceled events correctly in the ICS files. You can manually add events that have been cancelled to a spreadsheet to remove them from your calendar. Create a Google spreadsheet named `facebook_excluded_events` that the account running the sync can access. In column A of its first tab, enter one Facebook event UID (for example, `e123456789@facebook.com`) or event link (for example, `https://www.facebook.com/events/123456789/`) per row. A header is optional; blank and unrecognized cells are ignored.
