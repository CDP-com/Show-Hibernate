WScript.Echo(fnHibernateEnabled());

function fnHibernateEnabled()
{
   var WshShell = WScript.CreateObject ("WScript.Shell");
   var bKey = WshShell.RegRead ("HKLM\\SYSTEM\\CurrentControlSet\\Control\\Power\\");
   var retval = WshShell.RegRead("HKLM\\SYSTEM\\CurrentControlSet\\Control\\Power\\HibernateEnabled");
   return retval;
}

