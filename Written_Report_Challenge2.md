# Analysis of Stock Performance by Ticker

## 1 Overview of Project

This project was the second challenge in the Carleton University Business Analytics and Data Visualization Boot Camp.  Module 2 of the first Unit of the boot camp was intended to, through the use of VBA, teach the fundamental building blocks of programming languages. These skills included creating VBA macros, triggering pop-ups and inputs, reading and changing cell values, and formatting cells.  The project helped us develop our skills in using nested loops and conditionals to direct logic flow.  Writing pseudo code was a very helpful skill to develop, as the VBA scripts were detailed to complete the project objectives.

### 1.1 Purpose

The purpose of the specific project within this module was to assist a client, Steve.  He has been looking into and investment his parents made in a green energy stock, DAQO New Energy Corp(DQ).  After looking at DQ, Steve asked us to extend the analysis to a list of green energy stocks.  Concerned about how slow the program will run if he tries to analyze a larger number of stock.  We were provided with pseudo code for a possible refactoring of the original code and asked to analyze its performance.  Always wanting to strive for perfection, we will also research further code improvements.

### 1.2 Approach and Challenges

The analysis followed the general process looping through the rows of the stock transactions to find the first tranaction in a given year, for a given ticker and recording the years opening price.  The same process was followed to find the last transaction for the year's closing price. This was done for Steve earlier for only ticker DQ.  The DQAnalysis VBScript in the spreadsheet has been modified to provide Steve with a consolidated summary of DQ's rate of return.  Below is the code that allows us to analyze one sheet after another.  This was included in the workbook as a small demonstration for Steve of what could be done if he were to further engage my services.

    Dim DataYear(1) As String
    
    DataYear(0) = "2017"
    DataYear(1) = "2018"
    
    'Run analysis for every year of data
    For K = 0 To 1
                    
        Worksheets(DataYear(K)).Activate

The results for Steve were as follows:

| DAQO (Ticker: DQ) |                    |              |
| ----------------- | ------------------ | ------------ |
|                   |                    |              |
| Year              | Total Daily Volume | Return       |
| 2017              | 35796200           | 1.994458388  |
| 2018              | 107873900          | -0.626018859 |

The results are correct but not very pretty, although they are presented here better than they are in the Excel spreadsheet - but, Steve would probably like the spreadsheet to be improved (that could be future work for me).  In addition to improving the appearance, Steve wanted to see all of the green energy stocks for years 2017 and 2018, so another VBScript was created AllStocksAnalysis.  This new version of the VBScript was somewhat inefficient and clearly annoying because it refreshes the All Stocks Analysis worksheet as in loops though every ticker causing the screen to flicker like a strobe light.  You can see this in the code below, where j is the index for rows and i is the ticker index.

       Next j
       '6) Output data for current ticker
       Worksheets("All Stocks Analysis").Activate
       Cells(4 + i, 1).Value = ticker
       Cells(4 + i, 2).Value = totalVolume
       Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

   Next i

The refactor code, in VBScript AllStocksAnalysisRefactored, resolves this inefficency by moving the writing of these values to outside of the main row by row stock analysis loop as follows.

    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(i + 4, 1) = tickers(i)
        Cells(i + 4, 2) = tickerVolumes(i)
        Cells(i + 4, 3) = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i  

   Creating an array for the volumes and prices was done so that this refactoring could be achieved - the results were far superior for Steve.  In further researching the possible causes of slow macro enabled workbooks I came across a fabulous resource that I added to another VBScript called EvenFaster.

   I noticed that this Original code was constantly writing to the sheet so I found this code at:  
   https://www.dummies.com/software/microsoft-office/excel/10-ways-to-speed-up-your-macros/

  It states that Automatic Calculations and Screen Updating can be turn off to reduce interactions with the spreadsheet until they are needed.  
  Here is the code:

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    At the end of the subroutine you need to turn them on to complete the spreadsheet work.  Here is the code for that:
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    It makes an amazing difference, both in time and visual program performance.

## 2 Analysis and Observations
### 2.1 Analysis of 2017 Stocks

Looking at table 2.1 below you can see that this was a banner year for green energy.  Here are some observations that I had for Steve's consideration of the DAQO New Energy Corp (DQ) stock in comparison to other stocks in the same market:

- DQ was the highest preforming stock in 2017, significantly out performing the average stock  
- DQ was the least traded stock of comparable green energy stocks
- The highest trading stock SPWR had a positive return, however, it was one one of the poorest performing green stocks

![This is a sheenshot form stock the 2017 performance from VBA_Challenge.xlsm spreadsheet](Resources/StockAnalysis2017.png "Table 2.1 - 2017 Stock Analysis")  
**Table 2.1 - 2017 Stock Analysis"**

#### VBScript Performance on 2017 Ticker Data

The refactoring of the original VBScipt produced much better execution times as seen in the MsgBox screen shots below.  As fast as the refactoring was, you can see that it can still be improved.  Notice the progressive improvement in Figures 2.1.1, 2.1.2 and 2.1.3 for the 2017 data set.

![This is a MsgBox from the AllStocksAnalysis VBScript Module my VBA_Challenge.xlsm spreadsheet](Resources/VBA_Challenge_2017_before.png "Figure 2.1.1 - 2017 Stock Analysis VBScript Performance Before Refactoring")  
**Figure 2.1.1 - 2017 Stock Analysis VBScript Performance Before Refactoring**

![This is a MsgBox from the AllStocksAnalysisRefactored VBScript Module my VBA_Challenge.xlsm spreadsheet](Resources/VBA_Challenge_2017.png "Figure 2.1.2 - 2017 Stock Analysis VBScript Performance After Refactoring")  
**Figure 2.1.2 - 2017 Stock Analysis VBScript Performance After Refactoring**

![This is a MsgBox from the AllStocksEvenFaster VBScript Module my VBA_Challenge.xlsm spreadsheet](Resources/VBA_Challenge_2017-Even_Faster.png "Figure 2.1.3 - 2017 Stock Analysis VBScript Performance After Refactoring")  
**Figure 2.1.3 - 2017 Stock Analysis VBScript Performance After Refactoring to an Even Faster**

### 2.2 Analysis of 2018 Stocks

Looking at table 2.2 below you can see that, in stark contrast to 2017, 2018 was a very poor year for green energy.  Here are some observations I had for Steve's considerations of the DAQO New Energy Corp (DQ) stock in comparison to other stock in the same market:

- In 2018 DQ was now the lowest preforming stock of all of the green energy stock analyzed  
- DQ trade volume was up but still well below the average trading volume for this sector
- The highest trading stock in 2017 SPWR was still trading high, but the volumes had dropped significantly and they were now showing negative returns  
- Only two stocks, ENPH and RUN were showing positive returns in a market of big losers, their substantial positive returns were quite striking

![This is a sheenshot form stock the 2018 performance from VBA_Challenge.xlsm spreadsheet](Resources/StockAnalysis2018.png "Table 2.2 - 2018 Stock Analysis")  
**Table 2.2 - 2018 Stock Analysis"**

#### VBScript Performance on 2018 Ticker Data

The refactoring of the original VBScipt produced much better execution times as can be seen in the MsgBox screen shots below.  As fast as the refactoring was in you can see that it can still be improved.  Notice the progressive improvement in Figures 3.1.1, 3.1.2 and 3.1.3 for the 2018 data set.

![This is a MsgBox from the AllStocksAnalysis VBScript Module my VBA_Challenge.xlsm spreadsheet](Resources/VBA_Challenge_2018_before.png "Figure 3.1.1 - 2018 Stock Analysis VBScript Performance Before Refactoring")  
**Figure 3.1.1 - 2018 Stock Analysis VBScript Performance Before Refactoring**

![This is a MsgBox from the AllStocksAnalysisRefactored VBScript Module my VBA_Challenge.xlsm spreadsheet](Resources/VBA_Challenge_2018.png "Figure 3.1.2 - 2018 Stock Analysis VBScript Performance After Refactoring")  
**Figure 3.1.2 - 2018 Stock Analysis VBScript Performance After Refactoring**

![This is a MsgBox from the AllStocksEvenFaster VBScript Module my VBA_Challenge.xlsm spreadsheet](Resources/VBA_Challenge_2018-Even_Faster.png "Figure 3.1.3 - 2018 Stock Analysis VBScript Performance After Refactoring")  
**Figure 3.1.3 - 2018 Stock Analysis VBScript Performance After Refactoring to an Even Faster**

### 2.3 Overview of Analysis
#### What do we know?  
There really isn't much data to work with, but here is what we know:  
- What we do know is that the beginning of 2017 was a great time to invest in the green energy sector.  
- We also know that the end of 2017 and the beginning of 2018 would have been a terrible time to make the same investment. 

#### What don't we know?
There are an awful lot of things that we don't know.  In fact, there are far more things that are unknown than known and would be helpful if we knew:
- For one, we don't know how long each of these companies have been trading.  
- We also don't know when Steve's parents first made their investment - i.e. what was their original investment.
- We don't know anything at all about the market in general - i.e. was the big decline in 2018 market-wide or just in the green energy sector.  

#### How can we help?  
If Steve is willing to continue to pay for our services, we can offer Steve more assistance in answering some of these unknown questions.  Without knowing more than we currently know, it would not be advisable to try and draw any meaningful recommendations for Steve that would help his parents.

## 3 Challenges and Difficulties Encountered

### Additional Refactoring Example

I didn't like that the stock data had to be sorted before it could be analyzed, so for my own interest I created this script the gives the same results no matter how messed up the rows are.  It is pretty efficient too.  I didn't include this in the project, because I think we can 'up sell' Steve on some more work for large data sets.

        'I SET THIS LOOP UP SO THAT YOU DON'T CARE WHAT ORDER THE DATA IS IN IT WILL STILL WORK
        For i = rowStart To rowEnd
                    
                rowYear = Left(Cells(i, 2), 4)
                rowMonth = Mid(Cells(i, 2), 6, 2)
                rowDay = Right(Cells(i, 2), 2)
                
                tradedate = DateSerial(rowYear, rowMonth, rowDay)
                
                If Cells(i, 1) = "DQ" Then
                
                    totalVolume = totalVolume + Cells(i, 8)
                    totalTrades = totalTrades + 1
                    
                    If tradedate <= startingDate Then
                
                        startingDate = tradedate
                        startingPrice = Cells(i, 6)
                    
                    Else
                        endingDate = tradedate
                        endingPrice = Cells(i, 6)
                    
                End If
                End If
    
         Next i 

### Refactoring Pros & Cons
As mentioned at the introduction to this report there were a few data related issues that were effectively resolved though my work.  More importantly, there is a need for better, more complete, information - that is the quest of all consultants analyzing data.  This section is a summary and discussion of the Pros and Cons of refactoring in general and specifically in regards to Steve's requirements.

#### Pros of Refactoring
What has been done for Steve throughout this project has been of value because it facilitates his needs for future analysis on larger data sets.  The refactoring of code would really not have been of value if this was a one time analysis for Steve and this limited set of tickers was all that were to be reviewed.  In general terms, the refactoring of code is clearly of value as the number of data items  and/or the number of times the analysis will be run increases.  From my perspective, as an programmer, this exercise was of great value as I learnt a lot of tricks that I will use in the future.  I did notice that the original code was not designed well to make use of arrays and instead used variables that forced writing to the spreadsheet on every iteration of the tickers.  

Business cases for refactoring include:

- Clients who are willing to pay to have the work performed
- Businesses that will repeatedly save time because the code is part of their internal operations 
- Businesses that occasionally do analysis on very large sets of data - for example, monthly runs to assist in strategic investment or other business decisions 

#### Cons of Refactoring

From a professional learning perspective refactoring of code is always of value, and is something that developers can do for each other's code  - however as a businessman, there are clear situations where there are some definite cons.

Business cases for not refactoring include:

- If the application is a onetime initiative as the results have already been achieved - no need to spend more money, move on to a higher priority project  
- The application is conceptually very complex and the benefit of refactoring in terms of time saved could never outweigh the programming effort to speed up the process
- As you can see in the bullets above, in general, it is always a business case to determine whether refactoring will pay for itself


## 4 Results Summary and Recommendations
I have discussed a few observations above in the report, but the overall observations are as follows:
- Steve now has a valuable tool to help him analyze multiple stocks
- Steve is now aware of the current limitations of his data set
- Steve is also aware that we can modify the code to permit the analysis of unstructured data sets
- He also knows that if the data exists, we can easily add sheets for different years into the analysis including the ability to select tickers from a list to focus on multi-year analysis of promising investments

I would recommend that Steve further engage my consulting services to determine if, through the analysis of more information, better investment opportunities may exist in the green energy or other sectors of interest to his parents.
