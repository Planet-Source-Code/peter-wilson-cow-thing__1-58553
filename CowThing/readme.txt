' ================================================================================
' Cow Thing
' Version: 1.0
'
' by Peter Wilson
' Copyright © 2005 - Peter Wilson - All rights reserved.
' http://dev.midar.com/
'
' ================================================================================


Overview
=========
	Watch 28 frames of animated cows walk all over your desktop and windows, munch on the grass and moo occasionally. I was inspired by a web site I found whilst researching another project. I downloaded the cow animation and thought I might as well have some fun with it. The code is a little messy because I just threw it together on a lazy Saturday night. If someone would like to clean this up, add more animals and use polymorphism that would be great.


Features
========
	* Change the number of cow, with the following line:
		Private m_Cows(1000) As Cow

	* Transparency and/or Blue-Screen effects require Windows 2000 minimum.
	For a really cool effect try changing "bTrans = 255" to "bTrans = 128". It's in the DoModifyFormProperties() routine.

	* I load the images at run-time from the filesystem. If you embed the images onto the form, then VB converts them to bitmaps and the filesize increases.

	* Sorting the cows isn't really required, but it makes them look better when they are overlapping each other. I draw hem from bottom to top.

	* The sort routine probably deserves it's own submission. It's fast... very fast.


Helpful Links
=============
	I found the cow animation here:
	* http://www.reinerstileset.4players.de:1059/englisch.htm
 	If you are into RPG and need free graphics, then you need to visit this site NOW!

	I got inspiration for my sort routine here:
	* http://www.cs.ubc.ca/spider/harrison/Java/sorting-demo.html
	(The sort page has java applets and may take a while to load - if you're into sorting, then it's worth it)


Planet Source Code
==================
Here is a list of my other submissions. Most of them are well rated.
	*  A 2D Asteroids Game
	*  A 2D DotProduct Demonstration
	*  A 2D game - Froggies, a game of leap frog.
	*  A 2D Rotation Demo using SIN() and COS()
	*  A 2D Rotation Demo v2.0
	*  A 2D Rotation Lesson - Fly a UFO
	*  A 3D Lesson v2, Very Simple
	*  A 3D Lesson v3.1, Moderate
	*  A 3D Lesson v4, Advanced
	*  3D Studio v6.0 beta
	*  A collision avoidance system for games using DotProduct.
	*  A Matrix Multiplication Lesson using the game Asteroids
	*  A Simple Solar System Simulator, v1.0
	*  A Vacuum Fluorescent Display Simulator v1.0
	*  Asteroid Collisions (using the DotProduct)
	*  Convert Fonts to Vector Graphics using GetGlyphOutline
	*  RGB Colour Wheel
	*  TechniColor Mouse Trails
	*  TechniColor Mouse Trails v2
	*  LED Clock, aka. The LED Clock Challenge.
	*  Plus a few extras...



peter@midar.com
http://dev.midar.com/


