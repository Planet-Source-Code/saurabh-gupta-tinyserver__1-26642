Tiny Server v1.1.5
==================

By       : Saurabh Gupta
	   -------------
Web Page : http://connect.to/tinyserver
           ----------------------------


What is TinyServer?
-------------------
TinyServer is a very basic http server. This server can accept multiple requests at once. The server is only 60 kb. The default page, webpage directory and port number can be configured. The message window provides details of connections and errors if any. The server has been configured to accept a maximum of 100 connections. I have used the Winsock control in VB.  TinyServer, as of now, supports only the GET request. It also does not support any server side processing. The server can be used for testing websites on a local network before uploading to the Internet. TinyServer is open source and you can freely distribute it.


Using TinyServer:
-----------------

Click on Start button to start TinyServer. It will listen for connections at the specified port. The details of connections can be seen in the Message Box. TinyServer can be minimized to the system tray. After starting TinyServer type 127.0.0.1 or localhost in your browser address bar to see your website. If port has configured to a value other than 80 type 127.0.0.1:x or localhost:x where x is the port number.


Configuration:
--------------

Click the configure button to configure TinyServer. You can choose the directory in which your web pages are placed and the default page. You can also select the port at which you want TinyServer to listen (default: 80). When you start TinyServer for the first time the following defaults are loaded.
Webpage Directory = [The directory to which TinyServer was installed]
Default Page      = index.htm
Port              = 80	


Licensing:
----------

TinyServer is open source. That means you can freely modify its source code and use it in your own programs. But if you use it somewhere a mention would be nice ;-}

PS: I am not including the compiled exe with release. In case you do not have visual studio on your comp write to me for the exe.

Contact:
--------

This is the first version I am releasing, so I know there will be a lot of bugs. Feel free to write to me at saurabh_gupta@india.com for bug reports and suggestions.

Or visit the website:
http://saurabhonline.org


The Latest Version
------------------

Details of the latest version can be found on the TinyServer web page at
http://connect.to/tinyserver


Features to be added future versions:
--------------------------------------
* CGI Support
* POST support
* Remote administration (Show or hide server, change config etc)
* Request Logging

Improvements since version 1.0.5:
---------------------------------
- The configuration section is now working. The port can be configured.
- Windows API functions are used for reading files instead of VB functions, making it much     faster.
- Tiny Server can now be minimized to the system tray.
