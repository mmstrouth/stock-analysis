# Stock Analysis

## Overview

Initially, my friend Steve asked for a sheet that analyzed one particular stock, DAQO New Energy Corporation, so that he could advise his parents on investing their money. The analysis covered a total daily volume and yearly return. The first analysis revealed that DAQO was not the best option for Steve's parents to invest in. He then requested a way to analyze other stocks at the touch of a button. Steve is so happy with the anaylsis he has been provided thus far, he has requested a way to expand the dataset. Ultimately the purpose of this project was to use the process of refactoring to make the initial VBA script more efficient. 

## Results
Overall, refactoring the code saved time running the analysis, as shown in screenshot 1 and screenshot 2.

Screenshot 1

![This is an image](https://github.com/mmstrouth/stock-analysis/blob/d3649e15d3ec0d22aa0506fb6bdc3cc12061ca8f/vba_challenge_2018_module_time.png)

Screenshot 2

![This is an image](https://github.com/mmstrouth/stock-analysis/blob/d3649e15d3ec0d22aa0506fb6bdc3cc12061ca8f/vba_challenge_2018_refactored_time.png)

Unforunately the directions provided by the challenge were nons-sensical, asking for the 2018 and 2018, as shown in the screenshot below. I interpreted that to mean show the original time and the refactored time. 

![This is an image](https://github.com/mmstrouth/stock-analysis/blob/d3649e15d3ec0d22aa0506fb6bdc3cc12061ca8f/vba_challenge_direction_confusion.png)

The code itself required that an array configuration using for loops, a skill not taught or practiced in the module. I originally ran it without specifically telling it to record the results on the All Stocks Analysis Worksheet, which is where my screenshot came from. When I added the following code, indicating where the results should show up, I consistently received an error. 

![This is an image](https://github.com/mmstrouth/stock-analysis/blob/b2a20d2104413a45d843367e24739deb93e633c8/vba_challenge_worksheet_code.png)


## Summary

### Advantages and Disadvatages of Refactoring Code

Refactoring has a variety of advantages and disadvatages. The ultimate advantage is to make the VBA script more efficient. It can make the macro run faster as well as make the code itself easier to read and understand. The disadvatage is that it may take additional work up-front, after an acceptable script has already been created. In order to refactor, one must approach the script with a critical eye and determine ways to make it more efficient. Although the time may be well-spent in the long run, the task itself may feel tedious and odious in the short run. 


### Pros and Cons Applied to Original VBA Script

The overall pro to refactoring the original VBA script was the decrease in run time. The con was my frustration with the process; the amount of time it took me to make the code run did not feel worth it for just my friend Steve in the long run. 
