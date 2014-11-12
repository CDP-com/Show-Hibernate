// JavaScript library

/*-----------------------------------------------------*/
/*------------- Environmental Functions ---------------*/
/*-----------------------------------------------------*/

function fnToLocalURL( strEscaped )
{
    //Check for double escape
    if ( strEscaped.indexOf( "%2520" ) > -1 )
    { 
        strEscaped = unescape( strEscaped );
    }
    
    var sReturn = "";
    if ( strEscaped.indexOf( 'file' ) == 0 )
    {
        sReturn += strEscaped;
    } else {
        sReturn = 'file:///';
        sReturn += strEscaped;
    }

    var i;
    for (i=0; i<sReturn.length; i++)
    {
        sReturn = sReturn.replace(/%3A/i, ':');
        sReturn = sReturn.replace(/%5C/i, '/');
        sReturn = sReturn.replace(/%22/i, '');
        sReturn = sReturn.replace(/%5B/i, '[');
        sReturn = sReturn.replace(/%5D/i, ']');
    }
    var nPos1 = sReturn.lastIndexOf( "//" );
    if ( nPos1 > 18 )
    {
        var sReturn1 = ( sReturn.substr( 0, nPos1 ) + sReturn.substring( nPos1 + 1 ) );
        sReturn = sReturn1;
        //WScript.Echo(sReturn);
    }

    return sReturn;
}

function fnToLocalPath( strEscaped )
{
    var sReturn = "";

    var localUrl = strEscaped;
    var i;
    for ( i=0; i<2; i++ )
    {
        if ( localUrl.indexOf( "%" ) > -1 )
        {
            localUrl = unescape( localUrl );
        }
    }
    var n1 = localUrl.indexOf( "file:///" );
    if ( n1 > -1 )
    {
        sReturn = localUrl.substring( n1 + 8 );
    } else {
        sReturn = localUrl;
    }
    while ( sReturn.indexOf("/") > -1 )
    {
        sReturn = sReturn.replace("/", "\\");
    }
    sReturn = escape( sReturn );
    
    return sReturn;
}

// fnGetOs("name" or "version")
function fnGetOs(preturn)
{
   var retval= "";
   var nameOfOs = "";
   //HKLM "SOFTWARE\Microsoft\Windows NT\CurrentVersion" CurrentVersion
   var osVersion = fnReadReg( "HKLM\\Software\\Microsoft\\Windows NT\\CurrentVersion\\CurrentVersion" );
   if (osVersion.length > 0)
   {
      if (osVersion == "5.0") nameOfOs = "Windows 2000";
      if (osVersion == "5.1") nameOfOs = "Windows XP";
      if (osVersion == "5.2") nameOfOs = "Windows XP";
      if (osVersion == "6.0") nameOfOs = "Windows Vista";
      if (osVersion == "6.1") nameOfOs = "Windows 7";
      if (osVersion == "6.2") nameOfOs = "Windows 8";
      if (osVersion == "6.3") nameOfOs = "Windows 8.1";
      if (osVersion == "6.4") nameOfOs = "Windows 10";
   }
   
   if ( preturn.toLowerCase().indexOf( "name" ) > -1 )
   {
      retval = nameOfOs;
   }
   else
   {
      retval = osVersion;
   }
   
   return retval;
}

/*-----------------------------------------------------*/
/*------------------ File Functions -------------------*/
/*-----------------------------------------------------*/

function fnGetName( psFilePath )
{
    var shell = new ActiveXObject ( "WScript.Shell" );
    var sPathFilename = shell.ExpandEnvironmentStrings( psFilePath );

    var fso, afile, sName;
    fso = new ActiveXObject( "Scripting.FileSystemObject" );
    sName = fso.GetFileName( psFilePath );
    
    fso = null;
    shell = null;

    return sName;
}

function fnFileExists( psFilename )
{
    var lSuccess;

    var shell = new ActiveXObject ( "WScript.Shell" );
    var sPathFilename = shell.ExpandEnvironmentStrings( psFilename );
    var fso;
    fso = new ActiveXObject( "Scripting.FileSystemObject" );
    if ( ! fso.FileExists( sPathFilename ) )
    {
        lSuccess = false;
    }
    else
    {
        lSuccess = true;
    }

    fso = null;
    shell = null;

    return lSuccess;
}

function fnDeleteFile( pFilePath )
{
    var shell = new ActiveXObject ( "WScript.Shell" );
    var sPathFilename = shell.ExpandEnvironmentStrings( pFilePath );
    var fso, afile, nSuccess;
    try
    {
        fso = new ActiveXObject( "Scripting.FileSystemObject" );
        afile = fso.GetFile( pFilePath );
        afile.Delete( true );

        nSuccess = 1;
    }
    catch ( ed1 ) {
        nSuccess = 0;
    }
    finally
    {
      fso = null;
      shell = null;
    }
    
    return nSuccess;
}

function fnMoveFile( pFilePath, pTargetDir )
{
    var shell = new ActiveXObject ( "WScript.Shell" );
    var sPathFilename = shell.ExpandEnvironmentStrings( pFilePath );
    var sTargetDir = shell.ExpandEnvironmentStrings( pTargetDir );
    var fso, afile, nSuccess;
    try
    {
        fso = new ActiveXObject( "Scripting.FileSystemObject" );
        afile = fso.GetFile( pFilePath );
        afile.Move( pTargetDir );

        nSuccess = 1;
    }
    catch ( ed1 ) {
        nSuccess = 0;
    }
    finally
    {
      fso = null;
      shell = null;
    }

    return nSuccess;
}

function fnWriteLog( log_txt, io_mode )
{
    if ( lDebugMode )
    {
        if ( io_mode == undefined ) io_mode = 8;
        for ( var i=0; i<3; i++ )
        {
            try
            {
                var fso = new ActiveXObject( "Scripting.FileSystemObject" );
                var objLogTextStream = fso.OpenTextFile( fileLog, io_mode, true, 0 );
                objLogTextStream.writeline( log_txt );
                objLogTextStream.close();
                i = 3;
            }
            catch ( e7 )
            {
                //
            }
            finally
            {
               fso = null;
            }
        }
    }
}

function fnWriteFile( file_name, out_txt, io_mode )
{
    var i;
    for ( i=0; i<3; i++ )
    {
        try
        {
            var textOut = out_txt;
            var fso = new ActiveXObject( "Scripting.FileSystemObject" );
            var t = fso.OpenTextFile( file_name, io_mode, true, 0 );
            t.WriteLine( textOut );
            t.close();
            i = 3;
        }
        catch ( e7 )
        {
            //WScript.Echo(e7.message);
        }
        finally
        {
            fso = null;
        }
    }
}

function fnReadFile( file_name )
{
    var returnArray = false;
    try
    {
        var fso = new ActiveXObject( "Scripting.FileSystemObject" );
        var t = fso.OpenTextFile( file_name, 1, false, 0 );

        while (! t.AtEndOfStream)
        {     
            if (!returnArray) { returnArray = [] }        
            returnArray.push( t.ReadLine() );      
        }

        t.close();
    }
    catch ( e7 )
    {
        //WScript.Echo(e7.message);
    }
    finally
    {
       var fso = null;
    }
    return returnArray;
}

/*-----------------------------------------------------*/
/*----------------- Folder Functions ------------------*/
/*-----------------------------------------------------*/
function fnFolderExists( psFoldername )
{
    var lSuccess;

    var shell = new ActiveXObject ( "WScript.Shell" );
    var sPathFoldername = shell.ExpandEnvironmentStrings( unescape( psFoldername ) );

    var fso;
    fso = new ActiveXObject( "Scripting.FileSystemObject" );
    if ( ! fso.FolderExists( sPathFoldername ) )
    {
        lSuccess = false;
    }
    else
    {
        lSuccess = true;
    }

    fso = null;
    shell = null;

    return lSuccess;
}

function fnCreateNewFolder( psFoldername )
{
    var lSuccess;
    var shell = new ActiveXObject ( "WScript.Shell" );
    var sPathFoldername = shell.ExpandEnvironmentStrings( unescape( psFoldername ) );

    var fso;
    fso = new ActiveXObject( "Scripting.FileSystemObject" );

    if ( ! fso.FolderExists( sPathFoldername ) )
    {
        fso.CreateFolder( sPathFoldername );
    }
    if ( ! fso.FolderExists( sPathFoldername ) )
    {
        lSuccess = false;
    }
    else
    {
        lSuccess = true;
    }

    fso = null;
    shell = null;

    return lSuccess;
}



/*-----------------------------------------------------*/
/*--------------- Registry Functions ------------------*/
/*-----------------------------------------------------*/

// fnWriteReg(regkey, regvalue, regtype) returns true/false for success/failure
// Example: fnWriteReg("HKLM\\Software\\Myentry\\myname", "myvalue", "REG_SZ");
// regtype: "REG_SZ", "REG_EXPAND_SZ", "REG_DWORD", "REG_BINARY"
function fnWriteReg(regkey, regvalue, regtype)
{
   //WshShell.RegWrite ("HKLM\\Software\\Myentry\\myname", "myvalue", "REG_SZ");
   var success = true;
   try
   {
      var WshShell = new ActiveXObject("WScript.Shell");
      WshShell.RegWrite(regkey, regvalue, regtype);
   }
   catch (fWR)
   {
      success = false;
      //alert(fWR.message);
   }
   finally
   {
      WshShell = null;
   }
   
   return success;
}

// fnReadReg(regkey) returns the value of the setting or null if not read or it does not exist
// You can use this for testing if 
function fnReadReg(regkey)
{
   var regvalue = null;
   try
   {
      var WshShell = new ActiveXObject("WScript.Shell");
      regvalue = WshShell.RegRead(regkey);
   }
   catch (fRR)
   {
      regvalue = null;
      //alert(fRR.message);
      //null value returned in regvalue variable indicates that reg value was not read
   }
   finally
   {
      //destroy object
      WshShell = null;
   }

   return regvalue;
}

function fnDeleteReg(regkey)
{
   var success = true;
   try
   {
      var WshShell = new ActiveXObject("WScript.Shell");
      //var WshShell = WScript.CreateObject ("WScript.Shell");
      regvalue = WshShell.RegDelete(regkey);
   }
   catch (fDR)
   {
      success = false;
      //WScript.Echo(fDR.message);
   }
   finally
   {
      WshShell = null;
   }
   
   return success;
}

function fnRegExists(regkey)
{
   var retval = false;
   try
   {
      var WshShell = new ActiveXObject("WScript.Shell");
      WshShell.RegRead(regkey);
      retval = true;
   }
   catch (fRR)
   {
      retval = false;
      //alert(fRR.message);
      //null value returned in regvalue variable indicates that reg value was not read
   }
   finally
   {
      //destroy object
      WshShell = null;
   }

   return retval;
}



/*-----------------------------------------------------*/
/*------------------- Misc Functions ------------------*/
/*-----------------------------------------------------*/

// Search for string in an array returns array
// Example:
// if( !aArrayToSearch.find( unescape( sSearchString ) )
//    aArrayToSearch.push( unescape( sSearchString ) );
Array.prototype.find = function(searchStr) 
{  
    var returnArray = false;  
    for (i=0; i<this.length; i++) 
    {    
        if (typeof(searchStr) == 'function') 
        {      
            if (searchStr.test(this[i])) 
            {        
                if (!returnArray) { returnArray = [] }        
                returnArray.push(i);      
            }    
        } else {      
            if (this[i]===searchStr) 
            {        
                if (!returnArray) { returnArray = [] }        
                returnArray.push(i);      
            }    
        }  
    }  
    return returnArray;
}

