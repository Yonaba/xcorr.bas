__xcorr.bas__ is a small VBA module which adds up functions to MS Excel for [cross-correlation](https://en.wikipedia.org/wiki/Cross-correlation) computation.

##Installation##
Inside the Excel file where you need to use `xcorr.bas`, open *VBA* using `Alt+F11` combination, then add the module using `File > Import a File` dialog.

##Functions
* `CROSSCORRELATION (s1, s2, h)` where `s1` and `s2` are two signals and `h` a lag-time value
* `AUTOCORRELATION (s, h)` where `s` is a signal and `h` the lag-time value

##Credits and Thanks
I have been writing this small module for [Aïda Zaré](https://www.linkedin.com/in/zare-a%C3%AFda-97a50219/en), while she was trying some random theory during her [PhD thesis](http://www.2ie-edu.org/soutenance-de-these-de-aida-zare/) at [2iE](http://www.2ie-edu.org). I am releasing it to the community in case anyone finds it useful.

##License##
This work is under [MIT-LICENSE](http://www.opensource.org/licenses/mit-license.php).<br/>
Copyright (c) 2014-2016 Roland Yonaba