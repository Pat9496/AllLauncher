# AllLauncher
AllLauncher is a command line based launcher for games. It is designed to be used with some other frontend and provides automation options.

Start games by using:   `AllLauncher.exe <System> <ROM or Windows shortcut>`
Start frontend by using:   `AllLauncher.exe -Menu`


### Features of AllLauncher:

- **Close blacklisted processes
Processes in `ProcessBlacklist.txt` will be terminated before the game starts. Freeing up ressources for your game. 

- **Stop blacklisted services
Services in `ServiceBlacklist.txt` will be stopped before the game starts. Giving more ressources to the game and keeping unneeded distractions away.

- **Restart whitelisted services
Services in `ServiceWhitelist.txt` will be restarted after the game. Handy for services you still require outside of your game (such as printer spooler, etc.). 

- **Close and restart your Launcher/Frontend of choice
Some launchers or frontends look incredible but use up a lot of ressources. AllLauncher can quit your launcher and restart it after the game - freeing lots of ressources. AllLauncher itself is tiny. 

- **Open game related documents
Choose a directory and put your game related PDFs in there. Just name it like the ROM/shortcut and it will automatically be opened with the game - and closed after the game is finished. 
You need more than one document? Simply create a subdirectory and name it like the game and AllLauncher will open them all for you (and close them again). 

- **Automatically start a trainer
If you want to use a trainer for a game, you just need to name it like the ROM/shortcut and put it into your trainer directory. It will automatically be started with the game. 

- **Load Cheat Engine tables
If you want to use Cheat Engine tables, you can name them like the game and put it in your desired folder. AllLauncher will automatically start Cheat Engine and load the table. 

- **Load ArtMoney table
Just like Cheat Engine, ArtMoney does basically the same thing - with one exception: It can also load tables for various emulators, allowing you to make subfolders for each of your emulated systems, where you can put your tables in. 

- **Start additional programs
Does your game require come other tools? Just copy appropriate shortcuts into the games document folder and these they will be automatically started before the game (and closed again).

- **Open Universal Hint System file
[UHS](http://www.uhs-hints.com/) is a great tool for adventures and RPGs that only shows you the hints you actually need without poiling the rest of the game. If you copy a UHS-File with the game's name, it will be automatically opened.

- **Help your PS4-gamepad-management
Are you using DS4Windows? The you probably have it set to XBox-360-gamepad-emulation. AllLauncher will make sure, DS4Windows has exclusive access to your pad and provides two files - `DS4SystemBlacklist.txt` and `DS4WindowsBlacklist.txt` - where you can choose systems (i.e. emulator) and windows games that do not require DS4Windows because they natively support the PS4-gamepad. 

- **Manage Steam and Origin usage
AllLauncher automatically shuts down Steam and/or Origin, if they are not required. And it's a soft shutdown, meaning it will allow Steam and Origin to finish syncs if they are in the middle of it. 

- **Remove desktop distractions
AllLauncher will close all open file explorer windows and minimize everything else to clean up distractions. AllLauncher even goes one step further and hides your desktop as long as your game is running (you can still access it anytime).

- **Partial implementation for [Borderless Gaming](https://github.com/Codeusa/Borderless-Gaming/releases)
Want to use Borderless Gaming? Just put your game-shortcut name into `BorderlessGamesList.txt` and it will be automatically started. 

- **Set default audio volume for windows and your emulators
Just put your desired default volume into the `AllLauncher.ini` and it will always set your default volume before and after a game. 
Additionally, you can define volume adjustments for different systems (i.e. emulators). 

- **Audio volume "night mode"
You can also define certain timeframe where you want all your volume settings adjusted. This, of course, is only used for your normal speakers - not your headphones or VR-headset. 

- **Completely load or unload VR drivers and client as needed
*Currently only tested with Oculus Rift!*
VR requires a lot of ressources (services, client, dashboard, home, ...) so it makes sense not to load everything whenever your system is not "VR". 

- **Start SteamVR when needed
Certain games require SteamVR even if it is not bought through Steam. AllLauncher checks if `steam_api.dll` is present in the game directory and will start SteamVR when found. 

- **Partial [Reshade](https://reshade.me/)-Updater
Tired of manually replacing Reshade-DLLs everytime there is an update? AllLauncher will do that for you. 
*Does currently not work for Steam-games!*

- **Mount CD/DVD-images
*Only ISO-files supported at the moment!*
Game requires one or multiple disc-image(s)? Just put it/them in the game directory and AllLauncher will automatically mount it/them - and unmount it/them again after the game. 

- **Keep complete track of game-time
AllLauncher will log the time for each game and organize it in neat CSV-files. It will track, when and how long you played each game.

- **Text-To-Speech-Feature
Let your computer tell you what's going on and also tell you, how long you played the game this session and in total. 
*AllLauncher will automatically activate and use Microsoft Eva (the Cortana voice) for this.*
