# Delete Older Emails
Use when you want to delete emails from specific folder (or folder structure), but you are too lazy to do it manually (or it takes too much time).

## Data to be replaced
`<EMAIL>` - Name of your main folder (which usually is an email, but might be customised)
`<FOLDER>` - A folder name under the Inbox
> In line 13 you can remove or add more of `.folders.Item("<FOLDER>")` pieces as long as it satisfies your folder structure. You can also rename Inbox there if your Inbox folder has different name.
`<NUM>` - Number of months back from now (for example, if you have March and want anything older than February you pick `-1`)
> In line 27 you can customize how old emails can be. Notice that in line 28 you are replacing just picked month and year (if you decide to use negative number smaller than the month you have right now, Date will automatically recalculate the year as well), but you can also modify the day using Day() method.
