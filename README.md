# Observatory-control-app
A simple controller to improve observatory risk
This is a Windows Form app to control roof, switches, telescope mount and make safety related actions related to weather conditions and avoiding clashes between roof and telescope.
Its operation is highly dependent on particular equipment. It's function is not a given, and the user MUST undertake their own testing to ensure that all potential collision conditions are accounted for.
For example, this is currently used with a Paramount MX, which requires homing before anything else. Before homing, it does not know where it is and it is possible to home a mount into a closed roof without
additional sensors and ASCOM mount status to identify the issue and shut off the power.

I have added graphics to later versions to display environmental conditions and source code is reasonably documented so that one may modify it to taste.
