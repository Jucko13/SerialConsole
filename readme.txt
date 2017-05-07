go with a admin console to the path where the program is unpacked.

enter the following command: regsvr32 MSCOMM32.OCX

hit enter and start the program

to enable the support for ACD you need to add the following to the platform.txt of arduino:


################################################################################################
#             Stuff for ACD (Auto Connect Disconnect) by the SerialConsole program             #
################################################################################################
recipe.hooks.postbuild.0.pattern=F:\Github\SerialConsole\ZebroMote.exe "CC {serial.port}"
recipe.hooks.deploy.preupload.pattern=F:\Github\SerialConsole\ZebroMote.exe "CC {serial.port}"
recipe.hooks.deploy.errorupload.pattern=F:\Github\SerialConsole\ZebroMote.exe "CC {serial.port}"
recipe.hooks.deploy.postupload.pattern=F:\Github\SerialConsole\ZebroMote.exe "OC {serial.port}"
################################################################################################



platform.txt can be found in the following places FOR VISUAL MICRO:
    C:\Users\<username>\AppData\Local\Arduino15\packages\arduino\hardware\avr\1.6.16
OR FOR ARDUINO:
	C:\Program Files (x86)\Arduino\hardware\arduino\avr



