# AllLauncher
AllLauncher is a command line based launcher for games. It is designed to be used with some other frontend and provides automation options.

Start games by using:   AllLauncher.exe <System> <ROM or Windows shortcut>
Start frontend by using:   AllLauncher.exe -Menu


Features of AllLauncher:

- Automatically close blacklisted processes
  Processes in BlacklistProcesses.txt will be terminated before the game starts. Freeing up ressources for your game. 

- Automatically stop blacklisted services
  Services in BlacklistServices.txt will be stopped before the game starts. Giving more ressources to the game and keeping unneeded distractions away.

- Automatically restart whitelisted services
  Services in WhitelistSercices.txt will be restarted after the game. Handy for services you still require outside of your game. 
