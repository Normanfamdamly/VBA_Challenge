# VBA_Challenge Module 2

##  **Overview of Project**

This challenge focused on the refractor of coding. Steve had us pull the green stock analysis together, calculating the Annual Volume and the Annual Return for each stock for 2017 and 2018. In the process we made the code livable, so it could grow with Steve as the years passed. Now it was time to see if we could make the code more efficient. Efficiency comes in many forms such as:

    -Less Lines of Codes
    -Utilization of Less Memory
    -Better Logic
    -Easier to Read

## Results of Refacting 

If we are solely evaluating the results of the code on the time it took to run we could say the code was successful. We also need to look at the perfomance of th stocks between 2017 and 2018 and what if any effecienies we gained through the refactoring.  Was the savings of milliseconds worth the additions of new arrays and the changing of the code?

### Timing Sped Up
The 2017 Stock Analysis went from .7578125 down to .1328125 seconds to run, which is approximately a decrease of .63 seconds. While the 2018 Stock Analysis went from .765625 down to .171875 seconds to run, which is approximately a decrease of .59 seconds. There was no real decernable difference waiting on the processing of the stock analysis, per my 12 year old son. 
 ![2017_GreenStocksTimer](2017_GreenStocksTimer.png)
  ![VBA_Challenge_2017](VBA_Challenge_2017.png)
   ![2018_GreenStocksTimer](2018_GreenStocksTimer.png)
    ![VBA_Challenge_2018](VBA_Challenge_2018.png)

### Stock Performance
You can tell from the screen shots that the stock performance stays the same for both code versions.  The only standout item to note is 2017 was a much better year for the green stocks than 2018, in the later year the environmental stocks did not fair well and most returned a negative annual return. Hindesight is always 20/20 and the stock market crash and a Adminstration that was pro Gas and pro Coal may have had an effect on these company stocks. 

### Codeing Results
The biggest differnece between the Green Stocks code and the VBA Challenge code was the creation of the arrays for tickerIndex, tickerVolumes, tickerStartingPrices, and tickerEndingPrices. An array is like a set of store shelves we can put our information in as it is calculated and grab when it is needed.  The arrays sped the process of the calculations up by storing the information in the background as the code ran the loop. 


    ''1a) Create a ticker Index 
    ' Need an array for tickerindex, make Single since got Error 6 again at calculation of Return
    
    Dim tickerIndex As Single
    tickerIndex = 0
    
    '1b) Create three output arrays
    ' 12 tickers so denoting my floors to keep track and in like with ticker string
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
       
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    'index is floors and the bottom floor is 0 and goes to 11
    'tickerIndex will need to be used in my equations going forward- remember
    
    For tickerIndex = 0 To 11
    
    tickerVolumes(tickerIndex) = 0
    
## Summary of VBA Challenge

### Advantages and Disadvantages to Refactoring

#### Disadvantages
- Duplications of code due to copy and paste errors
- Logic that may be out of place due to jumping around the code you are cleaning up
- Too many hand in kichen if you have multiple refactoring
- If code still functions the same, maybe diffcult to prove to customer the time and effort put into the work

#### Advantages
- Easier to read
- Efficency in processing
- Liveable code that allows the program to live beyond the one spreadsheet
- Formatting updates make it more presentable to the consumer

### My Disadvantages and Advantages with the Challenge
    
There were more disadvantages to this challege than there were advantages. That may be in part to the newness of the the entire refactoring process, but I found a large disadvantage was trying to flip between the VBA Challenge code we were provided and my Green Stocks VBA code. There was no process in which I could have both open in the deveoloper and flip and and forth and not loose track of whcih code I was looking at, or working on/in.  If you copied the code over to a ReadMe text file that was no help either since my notes would all turn black in my Green Stocks VBA code.  The other disaadvantage of this refactoring is it was not a code I could make a change run, then move on and make another change and run then move forward.  THis code required all teh arrays and loops to work, so I felt I could not test as I went along.

The advantage to refactoring is obvious is does make a more efficent code and it made my code much easier to read. I found my Green Stocks VBA code used WAY TOO much wite space which made it difficult to read when I was compaing to the two codes.  I found a more condensed code with notes that made sense to me was easier to work when I came back from break.  THrough I question if someone else was working my code, if my notes make sense to them.  I did like the refactoring to an extent, since I am not a linear thinker, I could jump from piece to piece to fix in any order i wanted.


