# ExcelWaves
## A basic timing diagrammer for MS Excel

![Basic waveform](https://github.com/hacksterous/ExcelWaves/blob/main/basic-waves.png "Basic waveform")

Both the draw commands and the waveform are embedded in Excel.
### Note
1. Column A contains the row number at which the waveform will be drawn
2. Column A should contain only numbers.
3. Signal specifications should start from Column C.
4. Column B should be used for signal names.
5. Press Ctrl + w to execute the waveform drawing macro.
6. Wave labels, arrows and descriptions can be added as in regular Excel spreadsheet
7. Data labels and spacer definition should be placed on the cell immediately below the waveform specification.
8. Data labels can only be added below 'D' and 'X' specifications

#### Codes
```
1. l: low signal value
2. h: high signal value
3. c: low to high clock transition
4. k: high to low clock transition
5. z: tristate signal value
6. D: Data start
7. d: Data continue
8. X: Unknown bus start
9. x: unknown bus continuing
10. lz: low to tristate signal transition
11. hz: high to tristate signal transition
12. zl: tristate to low signal transition
13. zh: tristate to high signal transition
14. Xl: Like X, unknown bus start, will be followed by l
15. Xh: Like X, unknown bus start, will be followed by h
16. lX: Like X, unknown bus start, previous was l
17. hX: Like X, unknown bus start, previous was h
18. xl: Like x, unknown bus continuing, will be followed by l
19. xh: Like x, nknown bus continuing, will be followed by h
20. //: signal spacing
```
