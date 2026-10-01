-- accountedge_graph_mailer — OfficeLink Processor replacement applet
--
-- AccountEdge writes "OfficeLink Data" then launches this applet. The same
-- applet serves several jobs (email, an "are you there?" probe, Excel/Word
-- export). This replacement hands off to dispatch.sh, which sends EMAIL via
-- Microsoft Graph and passes every other job to the original applet unchanged.
--
-- The scripts own all user feedback (silent on success, a dialog on failure).
-- We wrap the call in try so the applet itself never throws a second dialog.

on aegm_run()
	set toolDir to (POSIX path of (path to home folder)) & "Library/Application Support/accountedge_graph_mailer/"
	set dispatchScript to toolDir & "dispatch.sh"
	set cfg to toolDir & "config.yaml"
	try
		do shell script "/bin/bash " & quoted form of dispatchScript & " --config " & quoted form of cfg
	end try
end aegm_run

on run
	aegm_run()
end run

on open theItems
	aegm_run()
end open
