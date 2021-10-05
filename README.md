# Stock-Analysis
## Analysis on VBA Deliverable 1
  The purpose of this analysis is to compare the effectiveness of using the orginal VBA script in contrast with the refractored VBA script.
  Speed, the code used, and the process to produce the code will be examined. 
### Speed 
  The orginal script that was developed in the course work delivered the 2017 results in .6171875 seconds. 
  <img width="204" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/90878901/135952134-9ae288f1-bedd-4f98-ad55-fbb0e2da636e.png">
  The orginal script from the course work for 2018 delivered the results in .59375 seconds. 
  <img width="176" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/90878901/135952222-65603cc1-b15f-4762-b0fb-02f52eb09c7f.png">
These speeds are acceptiable for this review. The data only contained 3013 rows for each year. Speed was also impacted by the formating in the code. 
### Coding
-The orginal script coding was easier on multiple issues and data to reflect how to build the code was methodical. The course work did a good job of explaining the process and how to develop the final results. This sequential step by step process was possible to follow. 
-The refractoring piece was a new concept that built off of thecourse work. This became a very difficult low reward, process that was not effective. At the time of the writting of this READ ME the answers are still not fully developed.  
 ### Code Development Process
 Of course the work on canvas helped develop the orginal code. Other online sources were used as well. https://docs.microsoft.com/en-us/office/vba/api/overview/ was a site used to help expand the knowledge of the orginal work. 
 The challenge for me was on the refractoring challenge. 
  1a) Create a ticker Index- for this section of code I used the following script: 
    Dim tickerIndex As Integer. Which seemed correct
  '1b) Create three output arrays. I used the following:
     Dim tickerVolumes As Long
	   Dim tickerStartingPrices As Single
	   Dim tickerEndingPrices As Single
     These are logical choices based on the need to create arrays. I was not able to see how these were used. See the example below: 
     
