# Stock Analysis
##### Automating the analysis of multiple renewable energy stocks using VBA
---

## Overview
This report will analyze the data for 2018 and 2017 of 12 renewable energy stocks including Atlantica Yield, Canadian Solar, Daqo New Energy Group, Enphase Energy, First Solar, Hannon Armstrong Inc., Jinko Solar, SunRun, SolarEdge, SunPower, Terraform Power, and Vivint Solar. The goal is to automate the formatting and analysis using VBA in order to output the yearly return and total volume for each stock. After creating the first draft of working code (Module 1: draftCode), I decided to *refactor* it to come up with a **design pattern** (Module 2: VBA_Challenge) with improved code performance that could be used on any sotck data.

## Analysis 
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;*Please, refer to the VBA_Challenge.xlsm in the repo specifically to the All Stocks Analysis worksheet. If you want to take a look to the entire code, open VBA and select the module called VBA_Challenge. There is only one Macro called "AllStockAnalysisRefactored".*

We first created an array for all the stocks, so that each one of them could be addressed with an index. 

![Challenge2Exhibit4](https://user-images.githubusercontent.com/83378141/119200473-009cd300-ba5b-11eb-985a-21bee79ee2e2.png)

Moreover, in order to come up with an automated analyis of the stocks we used a nested for loop that went through all the data, starting from the second row to the last one, storing the necessary information into arrays (total volume, first closing price, and last closing price) already initiated at the beginning of the macro. 

![Challenge2Exhibit3](https://user-images.githubusercontent.com/83378141/119200245-8d935c80-ba5a-11eb-9071-b0d85923326f.png)

As we can see in the lines of code above, after the loop began going through the rows we needed to create a conditional that if the row had data for the first stock we were looking for (`If Cells(i, 1).Value = ticker Then`), then we would store all the sum of volumes for all the rows that were from that same stock by storing it in each index of the totalVolumes array `tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value`.


## Results
![Challenge2Exhibit2](https://user-images.githubusercontent.com/83378141/119198757-ddbcef80-ba57-11eb-86bc-91491bb47286.png)
![Challenge2Exhibit1](https://user-images.githubusercontent.com/83378141/119198611-9cc4db00-ba57-11eb-88e5-ef5b0d346ace.png)

As we can observe 
