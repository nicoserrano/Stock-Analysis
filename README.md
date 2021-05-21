# Stock Analysis
##### Automating the analysis of multiple renewable energy stocks using VBA
---

## Overview
This report will analyze the data for 2018 and 2017 of 12 renewable energy stocks including Atlantica Yield, Canadian Solar, Daqo New Energy Group, Enphase Energy, First Solar, Hannon Armstrong Inc., Jinko Solar, SunRun, SolarEdge, SunPower, Terraform Power, and Vivint Solar. The goal is to automate the formatting and analysis using VBA in order to output the yearly return and total volume for each stock. After creating the first draft of working code (Module 1: draftCode), I decided to *refactor* it to come up with a **design pattern** (Module 2: VBA_Challenge) with improved code performance that could be used on any sotck data.

## Analysis 
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;*Please, refer to the VBA_Challenge.xlsm in the repo specifically to the All Stocks Analysis worksheet. If you want to take a look to the entire code, open VBA and select the module called VBA_Challenge. There is only one Macro called "AllStockAnalysisRefactored".*

I first created an array for all the stocks, so that each one of them could be addressed with an index. 

![Challenge2Exhibit4](https://user-images.githubusercontent.com/83378141/119200473-009cd300-ba5b-11eb-985a-21bee79ee2e2.png)

Moreover, in order to come up with an automated analyis of the stocks I used a nested for loop that went through all the data, starting from the second row to the last one, storing the necessary information into arrays (total volume, first closing price, and last closing price) already initiated at the beginning of the macro. 

![Challenge2Exhibit3](https://user-images.githubusercontent.com/83378141/119200245-8d935c80-ba5a-11eb-9071-b0d85923326f.png)

As you can see in the lines of code above, while the program was looping through all the rows it began storing values into the arrays. Beginning with the total volumes for each stock, I used the command `If Cells(i, 1).Value = ticker Then` to check if the first column had the ticker we needed. If it did, I then did `tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value` to store its volume in column 8 and keep adding to it everytime it encountered a row that was of the same stock. 

Furthermore, to check if the row was the first row for that stock I now needed two conditionals `If Cells(i, 1).Value = ticker And Cells(i - 1, 1).Value <> ticker Then`. If these conditions applied, it then went on and set the value in the 6th column (closing price for Jan 1st of that year) to become the first closing price of the stock. Which I used later to calculate the return. 

And lastly, I needed the last closing price of that stock (Dec 31st). In order to do that now I needed two conditionals `If Cells(i, 1).Value = ticker And Cells(i + 1, 1).Value <> ticker Then`. I needed to know that the row was for that stock and that the next row was not from that same stock, checking that first column as well. After storing the value of the last closing price (`tickerEndingPrices(tickerIndex) = Cells(i, 6).Value`), I needed to increase the ticker index so that it went on to the next stock and repeated the entire process for each one of them. 

As the data I wanted to output was already saved into arrays, I only had to create a for loop to do it. Then, it was just a matter of formatting the data so that it showed a green highlight if the return was greater than 0 and red one if smaller than 0 to allow for a better interpretation of the findings.

![Challenge2Exhibit5](https://user-images.githubusercontent.com/83378141/119203477-1f9e6380-ba61-11eb-86d6-6a6fb2a7a83d.png)



## Results

![Challenge2Exhibit2](https://user-images.githubusercontent.com/83378141/119198757-ddbcef80-ba57-11eb-86bc-91491bb47286.png)
![Challenge2Exhibit1](https://user-images.githubusercontent.com/83378141/119198611-9cc4db00-ba57-11eb-88e5-ef5b0d346ace.png)

As it can be seen just by glancing through the findings, the 2017 performance was considerably better as there is a lot more green to be seen in the table. This means there are more returns that came out positive. This is specially true for stocks such as Daqo, that in 2017 came out with a return of 199.4% and in 2018 of -62.6%. Similarly, this same happened for most of the stocks including Atlantica Yield (ticker: AY), Canadian Solar (CSIQ), First Solar (FSLR), Hannon Armstrong Inc. (HASI), Jinko Solar (JKS), SolarEdge (SEDG), SunPower (SPWR), and Vivint Solar (VSLR). 

In conclusion, even though the vast majority of them performed well in 2017, most of them performed poorly in 2018. Definitely, the top picks would have been Enphase Energy (ENPH) or SunrRun (RUN), which were the only green energy stocks that gave high returns for both consecutive years. The best investment would have been to buy a lot of ENPH stock at the beginning of 2017 and hold it through 2018 as it had 129.5% and 81.9% return, respectively. Nevertheless, money could also have been made by buying a put option on TERP at the beginning of 2017 as it returned -7.2% and then in 2018 another -5%. 

### Execution Times and Code Performance


## Summary

