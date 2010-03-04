active_directory Cookbook
=========================

Introduction
------------

These examples assume you are using the `active_directory module <http://timgolden.me.uk/python/active_directory.html>`_
from this site. The following are examples of useful things that could be done with this module on win32 machines.
It hardly scratches the surface of AD, but that's probably as well.

The following examples, except where stated otherwise, all assume that you are connecting to the current machine.
To connect to a remote machine, simply specify the remote machine name in the AD constructor, and by the wonders
of DCOM, all should be well::

   import active_directory
   c = ad.ad ("other machine")


Examples
--------

Show the interface for the .Create method of a Win32_Process class
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

The wmi module tries to take the hard work out of WMI methods by querying the method for its in and out parameters,
accepting the in parameters as Python keyword params and returning the output parameters as an tuple return value.
The function which is masquerading as the WMI method has a __doc__ value which shows the input and return values.

::

  import wmi
  c = wmi.WMI ()

  print c.Win32_Process.Create


Run notepad, wait until it's closed and then show its text
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

..  note::
    This is an example of running a process and knowing when it's finished, not of manipulating text typed into
    Notepad. So I'm simply relying on the fact that I specify what file notepad should open and then examining the
    contents of that afterwards.

    This one won't work as shown on a remote machine because, for security reasons, processes started on a remote
    machine do not have an interface (ie you can't see them on the desktop). The most likely use for this sort of
    technique on a remote server to run a setup.exe and then, say, reboot once it's completed.

::

  import wmi
  c = wmi.WMI ()

  filename = r"c:\temp\temp.txt"
  process = c.Win32_Process
  process_id, result = process.Create (CommandLine="notepad.exe " + filename)
  watcher = c.watch_for (
    notification_type="Deletion",
    wmi_class="Win32_Process",
    delay_secs=1,
    ProcessId=process_id
  )

  watcher ()
  print "This is what you wrote:"
  print open (filename).read ()

