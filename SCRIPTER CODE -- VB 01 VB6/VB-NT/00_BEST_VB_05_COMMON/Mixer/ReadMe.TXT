---------------------------------------------------------------------------
          Title | Matt Player
           Date | July 2 / 2004
         Author | Matt Gillmore
   Author Email | SCO_STINKS@yahoo.com
---------------------------------------------------------------------------
I make no promises that this code will do what you want or even that it
will do what I meant it to do.  You use it at your own risk.
---------------------------------------------------------------------------


[About]
It's a simple audio spectrum analyzer.  It opens a waveform audio input device
for 16-bit mono input, gets chunks of audio, runs the FFT on them, and displays
the output in a little window.

The Visual Basic sourcecode is included.  It demonstrates an easy and fairly
fast way to do graphics double-buffering with a hidden picturebox, audio input,
FFT usage, etc., etc.

NOTE:  The visualization will not work unless either the Wave OUT Mixer, or CDPlayer
is selected in the volume control under the Recording Control.  If anyone knows how to 
programmatically set these settings let me know.

[Aknowledgments]
Thanks go to DeethSource author Murphy McCauley who credits
Don Cross for providing the original FFT code.


[Requirements]
 a soundcard, of course.
