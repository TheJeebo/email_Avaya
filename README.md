# email_Avaya
Check Avaya queues and send email update

Opens connection with Avaya (scrubbed the original server, username, and password), collects the queue report, and calculates to determine if emails need to be sent. If so then use the send_Email method. Always logout and disconnect from Avaya so instances don't stack up.
