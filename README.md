Voice Email
===========

![preview](https://github.com/adityadalal924/voice-email/raw/master/screenshots/Outlook.png)


Voice Email is an Outlook extension to reply to emails with recorded messages.  

Features
--------
  - Natively integrated into Outlook (Shows up as tab on ribbon in message window)
  - Attaches recorded audio as MP3 automatically
  - Can add in template text (ie. "Reply in attached file")
  - Unlimited pausing/resuming
  - User-defined template text
  - Two different template text messages (ie. business vs personal)
  - Lightweight (about 2MB installed)
  - Fast conversion to space-saving format (LAME MP3 encoder)

Development
-----------
  - Set up Visual Studio for Office extension development as shown [here](http://msdn.microsoft.com/library/vstudio/bb398242.aspx)
  - Clone or download this repository, and open the solution in the source folder.
  - Create a Test Certificate (Properties > Signing) to sign the extension, as shown [here](screenshots/DevSetup.png)
  - Set Outlook to open when debugging (Properties > Debug > Start external program > Outlook path)
  - Configure Publish settings as necessary
  - Add lame.exe and BasicAudio.dll to "path" on [target system](screenshots/OutputFolder.png)
  - Verify installation of addon in [Outlook](screenshots/OutlookOptions.png)

Credits
-------
   - [Basic Audio](http://basicaudio.codeplex.com/) for audio recording
   - [LAME MP3](http://lame.sourceforge.net/) for MP3 conversion
   - [Icons8](http://icons8.com) for application icons


License
--------
Voice Email is licensed under the [MIT License](https://www.tldrlegal.com/l/mit), copyright (c) 2014 Aditya Dalal
