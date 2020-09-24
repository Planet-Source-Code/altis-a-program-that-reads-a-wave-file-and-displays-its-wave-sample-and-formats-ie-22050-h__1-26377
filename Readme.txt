Copyright Notices: This file MUST be included when redistributing.  You may not remove or alter the Readme.txt file that was originally distributed with it.

**************************************************
Get the most out of wave file.

This is an update of the 'Read a wave file' code that I have submitted a couple of years ago.  The old code does not read a wave file in a correct way, and I feel that it does not provide enough helps.  Therefore, I used the correct way to read the wave file in this project.  This program demonstrates how to read a wave file and displays the wave samples (If there is more than one channel i.e. Stereo, it will display those two channels separately.), Samples rates per second (i.e. Hertz), Average bytes per second (How many bytes play every second), Bits per sample (Bit resolution of a sample point.  i.e. 16-bits) and the length of the wave file in terms of seconds.  The program also allows the user to enlarge a specific section of the wave file.  When the user plays the wave file, he or she will be able to see a line that indicates the byte that is being play.

**************************************************
Some of the useful Equations:
MRATIO, GRATIO
Ratio = (Height / (Number of Channels + 1)) / (Distance between the highest and lowest point of the wave sample)

Length of wave file in terms of Seconds
Length = (Size of Wave file / wBlockAlign) / Samples per Second

Twips to Number of Bits, X as twip, dBitsPerTwip as the number of bits drawn in one twip
X * (dBitsPerTwip * Number of channels) / wBlockAlign
Reverse is:
B * wBlockAlign / (dBitsPerTwip * Number of Channels)
	Where B is the number of bits.

**************************************************
For more information on wave formats, go to www.wotsit.org and search for 'Wave Formats'.
	OR
go to msdn.Microsoft.com and search for 'Reading Wave'.  (Recommend Article Four)
**************************************************
Visual Basic 5.0 Users:

Replace the CommonDialog control with the 5.0 version.  No rename require.