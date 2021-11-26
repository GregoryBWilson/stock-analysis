# Analysis of Stock Performance by Ticker
<!-- There is a title, and there are multiple paragraphs (2 pt).
Each paragraph has a heading (2 pt).
There are subheadings to break up text (2 pt).
Links are working, and images are formatted and displayed where appropriate (2 pt). -->

<!-- There is a title, and there are multiple paragraphs (2 pt).
Each paragraph has a heading (2 pt).
There are subheadings to break up text (2 pt).
Links are working, and images are formatted and displayed where appropriate (2 pt). -->

## 1 Overview of Project

This project was the second challenge in the Carleton University Business Analytics and Data Visualization Boot Camp.  Module 2 of the first Unit of the boot camp was intended to through th euse of VBA, learn the fundamental building blocks of programming languages. These skills included creating VBA macros, triggering pop-ups and inputs, reading and changing cell values, and format cells.  The project help us develop our skills in using nested loops and conditionals to direct logic flow.  Writing pseudo code was a very helpful skill to develop as the VBA scripts were detailed to complete the project objectives.

&&&&&&&

### 1.1 Purpose

<!-- Overview of Project: Explain the purpose of this analysis. -->


  
The purpose of the specific project within this module was to assist a client, Louise who came close to, but failed to meet her goals in funding a play named Fever.  It is not clear why Louise wants this information, but I am assuming that she probably wants to make another more focused attempt to fund this play.  As we strive for perfection, we will definitely be giving Louise the best information available.
### 1.2 Approach and Challenges
 
The analysis followed the general process of breaking the available data into categories and subcategories that were appropriate to Louise's needs.  She was interested in theatre and in particular plays.  I looked at two specific factors that may influence the goal outcome, those being launch date and the actual goal amount.  This data is presented in a line chart to show trending.  I also did an analysis by country and extracted three indicative candidates to discuss with Louise.  The country data is presented in two small tables, table x.1 and x.2.
  
I created a number of views to verify the analysis was correct and meaningful.  I did observe that there was an issue in the project defined requirements in that the Outcomes based on Goals last row said Greater than 50000 and the second last row said 45000 to 49999.  This did not affect the values because there were no goals of exactly 50000 for theatres - it would cause issues for other campaigns.  To correct this, I changed the goal to be Greater than 49999.  Another issue was that I wanted to give Louise better data about what she could do in the future, so I created a couple of pivot tables that allowed me to extract data that I specifically wanted to discuss with her.  The last item was that the first row was a much smaller range than most others.  The second row had a slightly smaller and the last row was of course much larger.  Using a line graph chart could be somewhat deceiving, however, the trend in meaningful zones was however valid and I have addressed this in the explanation.

## 2 Analysis and Observations
<!-- Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script. -->

### 2.1 Analysis of Outcomes Based on Launch Date
  
The analysis of the database included a significant number of projects in the theatre category and most of those were from the subcategory plays.  The campaign for the play Fever that was launched in June may have been slightly late.  In the graphs x.x and x.y below you can see that May appears to be the best time of year to have a successful campaign.  This is particularly true when you analyse the success and failure ratio - the decline from May to the end of the year is significant.  You will also notice that while the success rises from January to May it is not as steep of a rise and then you find that in May and June the largest numbers of campaigns are launched.  I would conclude the following:
- When you launch a campaign too early in the year it is possible that investors may be inclined to hold off to see what other opportunities may present themselves
- If you wait too late to launch, then you are likely to find that many investors have already committed their funds to another project  

![This is a graph from my Kickstarter_Challenge.xlsx spreadsheet](Resources/Theater_Outcomes_vs_Launch.png "Figure 2.1.1 - Theater Outcomes vs Launch Date - Counts by Subcategory")
**Figure 2.1.1 - Theater Outcomes vs Launch Date - Counts by Subcategory**
![This is a graph from my Kickstarter_Challenge.xlsx spreadsheet](Resources/Outcomes_Sucess_Fail_by_Month.png "Figure 2.1.2 - Theater Outcomes vs Launch Date - Success/Fail Ratio")
**Figure 2.1.2 - Theater Outcomes vs Launch Date - Success/Fail Ratio**
### 2.2 Analysis of Outcomes Based on Goals

The analysis of the database shows that Louise's campaign, at less than $3,000 was set at a very reasonable goal level.  The only goal level that performed better was at less than a $1,000.  However, if you consider that most ranges were $4,999 you could in fact say that Louise was in the highest success range.  You could also say that while the less than $1,000 range was the most successful at raising money, it was also somewhat of an outlier in that it was likely too small a goal to achieve anything of importance.  I would argue that Louise was well positioned in terms of goals to be successful.  There are very few campaigns with goals that exceed $25,000 and in fact it is unlikely that the results above this range are of any statistical significance based on the small sample size.  The conclusions that could be drawn are:
- There is a clear trend that demonstrates that the larger the goal the less likely it is to get sufficient pledges
- The other interesting observation is that there seems to be a significant financial threshold at the $5,000 goal level where investor's interest tends to wain

![This is a graph from my Kickstarter_Challenge.xlsx spreadsheet](Resources/Outcomes_vs_Goals.png "Figure 2.2.1 - Theater Outcomes vs Goals - Counts by Subcategory")
**Figure 2.2.1 - Theater Outcomes vs Goals - Counts by Subcategory**
### 2.3 Analysis of Outcomes Based on Country

**What do we know?**  
This is what we have determine based on the work that Louise has contracted us to do.
- Louise's campaign to seek investment in June was a little late in the year just missing the May peek season.  However that is typically still in a fairly successful time of the year for investor funding
- Louise's goal was reasonable in terms of what she was looking to raise as she was well below the $5,000 apparent psychological threshold
- Unfortunately, Louise fell short of her goal

**How can we help?**  
Not wanting to leave Louis without a more significant plan to improve her chances next time, I decide to have a look at the market in which she was competing.  Using a pivot table with countries as rows, I noticed that three countries were worth looking at for the theater/play category: Canada, Great Britain and the United States.  Tables 2.3.1 and 2.3.2 below yield some valuable information for Louis:
- While the United States has by far the most campaigns the overall success rate is only 62%
- In Great Britain the success rate is an impressive 77% overall with successes 3.4 times more likely than failures
- If Louise feels the logistics of a launch in Great Britain are too difficult, she might consider Canada as an option due to the favourable investment environment and the relative proximity to the US

**Table 2.3.1 - Theater Outcomes vs Goals Success Rate by Country**  
![This is a graph from my Kickstarter_Challenge.xlsx spreadsheet](Resources/Success_Rate.png " Table 2.3.1 - Theater Outcomes vs Goals Success Rate by Country")

**Table 2.3.2 - Theater Outcomes vs Goals - Success/Fail Ratio by Country**  
![This is a graph from my Kickstarter_Challenge.xlsx spreadsheet](Resources/Success_Fail_Ratio.png "Table 2.3.2 - Theater Outcomes vs Goals - Success/Fail Ratio by Country")


## 3 Challenges and Difficulties Encountered
<!-- Summary: In a summary statement, address the following questions.
What are the advantages or disadvantages of refactoring code?
How do these pros and cons apply to refactoring the original VBA script? --> 
As mentioned at the introduction to this report there were a few data related issues that were effectively resolved though my work.  More importantly, there is a need for better, more complete, information - that is the quest of all consultants analyzing data.  For example, if it is true that projects who set a low goal are most successful at the beginning, are they also the most likely to fail in production - we don't know because the data is not available. We also know nothing about who is funding the campaign, it would be valuable to know which investors are open to what Louis has to offer. Most projects are funded at a very low level, unfortunately there are not enough data points to reasonably zoom in on the detail of that range to look for opportunities for Louise. 
## 3 Results Summary and Recommendations
I have discussed a few observations above in the report, but the overall observations are as follows:
- Louise has a viable ask in terms of financial needs
- Louise needs to improve her timing so that she hits investors at exactly the right time
- Louise may wish to look at Great Britain or even Canada as more amenable markets for her play

I would recommend that Louise further engage my consulting services to determine if it would be feasible to launch in a geographic market other than the United States.
