# Breaker
Get away from the computer and take a break.

Breaker is an idea I had kicking around, as to how I could force myself to take a break, as I tend to forget time and place when I'm "in the zone". A simply reminder doesn't cut it for me, since I'll just ignore it or dismiess it -- it's not enough of an interruption. So I thought about just adding a simple PowerShell script as a scheduled task, which would lock my scren at specified times. I then thought this might be something others would want, so I had to add a bit of user friendlynes.

The approach was initially to have a small (PowerShell) program, that would get input from the user, then add the relevant scheduled task and calendar entry. I've not changed my mind on that approach, to something that is, I think, much simpler: Just create a regular appointment with some specific keywords. The scheduled PowerShell script then just needs to periodically check if any appointments exist with the specified keywords. Maybe I'll get around to finishing it, at some point.
