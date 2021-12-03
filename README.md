# Stock Analysis with Excel VBA

## Overview of Project

### Purpose
The task was to analyze 12 stocks from 2017 and 2018 and determine their total daily volume and yearly return for each stock. By looking at the total daily volume we will be able to get a snaphsot of how actively a particular stock is being traded. While the yearly return will allows us to measure how well the stock performed from the beginning of the year to the end.

## Results

### Stock Performances between 2017 and 2018

#### A Look at 2017 Total Daily Volume
When we analyze the stock for 2017, the results provide a variety of stocks that performed well with varying degree of total daily volume and return in the positve.

![Screen Shot 2021-12-03 at 9 36 44 AM](https://user-images.githubusercontent.com/93485455/144629677-23094509-76e4-49a0-aa7b-23b5c20cd5a6.png)


To get a better understanding of the stocks that were most active we sorted the results to show the most active at the top of the list.

![Screen Shot 2021-12-03 at 9 06 21 AM](https://user-images.githubusercontent.com/93485455/144629724-5dde953c-fd05-40f9-85e7-de45e2e289bf.png)

Keep in mind as the list shows several stocks that had been very active and had very positive returns, this is only sample size of years' worth of data we must continue on with our analyis before making any rushed conclusions.

#### A Look at 2017 Return
It was a good year for many of the stocks on the list. Only one stock, TERP finished the year below where it started at the beginning of 2017. Again we sorted the table to show the best performers for the year to display them at the top of the list.

![Screen Shot 2021-12-03 at 9 08 02 AM](https://user-images.githubusercontent.com/93485455/144631444-a3e65087-2b13-415b-bf33-9c452f85a8b0.png)

If 2017 data is all that we had, then our analysis would give us a few stocks that could possibly be investment candidates but with only the year sample it leaves the decision to invest a more difficult one to process. Thankfully we do have 2018 data to analyze and compare to 2017, so let us look at how these stocks faired in 2018 to see how we would have done.

#### A Look at 2018 Total Daily Volume
The story in 2018 seems to have a completely different one. 

![Screen Shot 2021-12-03 at 9 59 07 AM](https://user-images.githubusercontent.com/93485455/144633183-8bf22c6a-6cf6-4792-b7fc-d23aeb5b0d43.png)
 
At first glance we see that many of the Returns were in the negative but let's review the Total Daily Volume. Let's sort the table again by the largest volume at the top.

![Screen Shot 2021-12-03 at 9 05 39 AM](https://user-images.githubusercontent.com/93485455/144633407-36bac5d0-5202-4fbe-90b0-9390b0ff43ed.png)

It seems we are seeing the stocks having a different Total Daily Volume than 2017 but let's look at the glaring difference as we review the Return for 2018.

#### A Look at 2018 Return
![Screen Shot 2021-12-03 at 9 08 42 AM](https://user-images.githubusercontent.com/93485455/144648659-dba79a2a-3df7-4354-8ef0-f04d55eb9720.png)

We see that there are only two stocks that performed better overall for the year – RUN and ENPH. But let's now try to compare the two years and see what the data tells us.

#### Comparing 2017 to 2018
![Screen Shot 2021-12-03 at 12 34 40 PM](https://user-images.githubusercontent.com/93485455/144654937-cf708786-6147-4073-bc3e-9c5947ac428e.png)

To getting a better glimspe in comparing the two data sets we have combined the tables and have two additonal columns with some calculations based of what we know. By adding the column Difference in Total Daily Volume, where we take the volumes for 2017 and compare to the volumes of 2018, it is clear that there are seven stocks that more activity in 2018 than 2017. But to get an even clearer picture of what stocks did the best over the two years lets sort the Return Over 2 Years with the best results at the top. The stock ENPH had the best return over two years with 211.45% followed by SEDG, DQ and RUN.

### Performance of Original Script
The original script had the run times as follows:

Original 2017 Script Result

![Pre-refactor 2017 Results](https://user-images.githubusercontent.com/93485455/144658035-05d39329-28c7-4bf0-bf67-0155836c7038.png)

Original 2018 Script Result

![Pre-refactor 2018 Results](https://user-images.githubusercontent.com/93485455/144658076-c41f2b72-8139-4064-90c4-82a083a568c3.png)


### Performance of Refactored Script

Refactored 2017 Script Result

![VBA_Challenge_2017](https://user-images.githubusercontent.com/93485455/144658323-bade394b-a68c-426b-b4f6-1f28ddc0ee99.png)

Refactored 2018 Script Result

![VBA_Challenge_2018](https://user-images.githubusercontent.com/93485455/144658344-b3a31576-7485-436a-875e-73574f57eeff.png)


## Summary

- What are the advantages or disadvantages of refactoring code?

 > The purposes of refactoring according to Martin Fowler (Father of Code Smell) are stated in the following:

 > 1. Refactoring Improves the Design of Software
 > 2. Refactoring Makes Software Easier to Understand
 > 3. Refactoring Helps Finding Bugs
 > 4. Refactoring Helps Programming Faster

 It also allow the code to be more adaptive over time and allows for the ability to bring new developers into the code without much training if refactored properly.
 
 The disadvatanges of refactoring code is that it takes time and money which may not be available or limited. 

- How do these pros and cons apply to refactoring the original VBA script?

 The disadvantages just mentioned do not come into play in our execercise but all of the advantages do. Let's address them one at a time.
 
  1. **Refactoring Improves the Design of Software** in our exercise allowed us to make the code more streamline and allowed for the script to address if the data       set were to change or grow in size and the run time is greatly improved thus not taxing the memory of the system and making the user experience better.
  2. **Refactoring Makes Software Easier to Understand** in our exercise by reducing the amount of code and adding more concise commenting on the script, it allows     for us to revist this code later down the road or someone entirely new to review the code and have a good understanding of what is trying to be acheived.
  3. **Refactoring Helps Finding Bugs** which in our exercise there were a few that cropped up but because the code was paired down, it allowed us to find the           issues and resolve them quickly.
  4. **Refactoring Helps Programming Faster** because now that the code is streamlined and commented, it allows for us or anyone else to add functional code to         improve and expand the capabilities of the script.

