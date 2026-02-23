### BUGS ###

- When resizing the window, the logo box seems to be misscalculating the vertical center, and i see flashes of the box being redrawn above and below whare it should be. please note the hrizontal center in not seeing any issues. this behavior does not happen when expanding the window in either axis. but is consitnantly happening when shrinking the window on the horizonal axis.

- after resizing the window, frequently, but not every time, the button click detection stops working. it is fixed by pressing one of the hotkeys. we need to fix whats causing that, and additionaly add some kind of routine after window resize that ensures button click detection in the console is in a working state.

- there is an unwanted blank space after the hourglass in the header, for the "Current" time setting. its only present in debugmode, possibly due to the "- DEBUGMODE" being printed in the header


### IMPROVEMENTS ###

- create an animation of the stats box sliding into frame from the right side of the screen and sliding out, for when the full view is enabled and disabled.

- Need to improve the mouse movement logic so that we are calculateing the number of points on the graph to be generated, against the mouse movement speed. so that ever 5ms the mouse moves to its next point along the curve. The calculation of the mouses acceleration and deceleration, should be factored into where the points are being placed on the curve not the frequency at which the mouse pointer is moved, i want that to remain consistant. its ok if we need use the same cooridinates multiple times. as long as we never move backwards along the graph.

- i want to smooth out the curve a bit right now its rather intense.

- currently the mouse buttons are triggering on mouse down. i want them to trigger on mouse up. i want each button to have its own color options for both foregorund and background. and an onclick foregroud and backgroud. so that on mouse down we can redraw the button with onclick colors.
    - once that is implemented at a basic level we can tweak how that onclick color persists, idealy it will remain that color while its coresponding popup window is open. or in the case of hide output that color should perisist for the (h) buton on the hide_output screen to give a since of toggling on and off.

- need to add a loading/welcome screen that diplays during the initialization, when running defaults there will be a puase so that the screen does not auomatically close when the initialization is completed. allowing any information to be read. of any paramaters at all are passed to the module, we will not pause after initialization is completed.

- when resizing the window i want to add logic to where if the mouse is still being held down, when the wait timer after the resizing stops, we do not immidiatley exit the resize screen while the mouse is still held. if the timer has ended, the moment the mouse button goes up we exit the resize loop. The mouse being held, should never reset the wait timer, only window resizing. It should just be tracked as a seperate requirtemtent for exiting the resize window ui loop.

### STATS SECTION ###

- i want to look into how to draw graps in the stats section. newer versions of window terminal supports witing pixels to the console. i will want to leverage that. to draw the curves being applied to mouse movment, and graphs depecting statistics. those options should only appear, and be used if the browser supports it. this should be deteinied suring the initialization

### OTHER ###

- Getting started with publishing to PSGallery, via github pipeline  https://www.youtube.com/watch?v=TdWWUOJ4s7A

