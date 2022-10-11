# UdonMidiPersistence
Midi Persistence allows persitence in VRChat worlds by saving data locally to your PC. Data is transmitted out of VRChat via the log file & transmitted into vrchat usind Midi.

<b>Midi Persistence runs 100% offline</b>

# Setup
## Requirments
These need to be installed for Midi Persistence to work.
### Loop Midi
https://www.tobias-erichsen.de/software/loopmidi.html
Loop Midi is a windows application for creating virtual midi ports. Without it Midi Persistence is not able to communicate with VRChat.

After starting Loop Midi for the first time, click the plus button to <b>create a new virtual midi port. Keep the name the default name</b>.

I recommend enabling auto start using the task bar icon.

<img width="236" alt="image" src="https://user-images.githubusercontent.com/40315315/195083708-bb46c4b2-998b-4bfc-bef0-52c3bd19eb97.png">


### .Net Runtime
https://dotnet.microsoft.com/en-us/download
.Net is a framework by microsoft. If you have visual studio you probably have it already installed, else you need to install it from the above url.

## Install
[Download the newest release](https://github.com/Thryrallo/UdonMidiPersistence---App/releases) and run the installer.

I recommend enabling auto start using the task bar icon.

<img width="156" alt="image" src="https://user-images.githubusercontent.com/40315315/195083979-b5285954-b313-4554-9a28-ec88af8b5d37.png">

## Updating
To update Midi Persistence first uninstall the old version, then install the new version.
Your data will be kept between the versions.
(I had issue with files not being overwritten using the updater, have to look into updater in the future again.
