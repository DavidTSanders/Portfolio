###Purpose
BikeCycle is a google sheets based bicycle tracking system for FreeRide, a Do-It-Yourself bicycle cooperative located in Pittsburgh PA.

BikeCycle's purpose is to track the lifecycle of a bicycle through FreeRide from the time a bicycle is first donated, to when it leaves the shop.

###FreeRide
FreeRide accepts donated bicycles in any condition. Volunteers evaluate the bike, fix or donate them as-is to another organization, or take them apart for metal recycling. Once a bike is fixed, it's donated or sold to the public at a drastically reduced price.

## Current Spreadsheets
There are two speadsheets currently being used by FreeRide  to record bicycle information, **Incoming** and **Outgoing**. 
###Incoming spreadsheet:
FreeRide recently started numbering incoming bikes and recording those numbers using a form that is linked to a spreadsheet named Incoming. This spreadsheet captures basic information about the bicycle such as donation date, manufaturer, and various size infomation. An abbreviated example is below:

| *Incoming*           |       |                |            |             |
| ------------------ | ----- | -------------- | ---------- | ----------- |
| **Timestamp**      | **Make**  | **Model**      | **Wheel Size** | **Bike ID** |
| 7/29/2022 10:12:14 | Kent  | Ambush         | 20"        |  A0001      |
| 7/29/2022 10:12:54 | Huffy | Stone Mountain | 20"        |  A0002      |
| 7/29/2022 10:13:32 | Next  | Screamer II    | 20"        |  A0003      |

###Outgoing spreadsheet:
When a bicycle leaves the shop through scrap, donation or sales, the bike Id and destination are recorded using a form that populates a spreadsheet named Outgoing.

| *Outgoing*    |                       |
| ------------|---------------------- |
| **Bike ID** | **Outgoing Destination**  |
| A0069       | Volunteer            |
| A0301       | Sold                 |
| A0302       | Communicycle         |



## Proposed Spreadsheet
### BikeCycle spreadsheet:
I was asked to take the two existing spreadsheets, Incoming and Outgoing, and create a third spreadsheet named BikeCycle. The BikeCycle spreadsheet will link the Incoming spreadsheet to the Outgoing spreadsheet using the Bike ID.


| *BikeCycle* |||||||
| ----------- | -------------------- | ------------------ | ----- | -------------- | ---------- | ----------- |
| Bike ID | Outgoing Destination | Timestamp          | Make  | Model          | Wheel Size | Bike number |
| A0069       | Volunteer            | 7/29/2022 10:12:14 | Kent  | Ambush         | 20"        |  A0001      |
| A0069       | Volunteer            | 7/29/2022 10:12:54 | Huffy | Stone Mountain | 20"        |  A0002      |
| A0069       | Volunteer            | 7/29/2022 10:13:32 | Next  | Screamer II    | 20"        |  A0003      |

##The Challenges
There were a few challenges to overcome to complete this project. 

+ The two original sheets are in seperate files. 
+ The lookup column in the Incoming spreadsheet, Bike Id, is not the first column in the sheet. This presents challenges for the common spreadsheet formula VLOOKUP() which requires that the lookup field be to the left of any fields returned by the search.
+ There are other spreadsheet options that can be used, but this is a showcase of my VLOOKUP() skills. Other options include:
    + INDEX() / MATCH()
    + QUERY()
    + XLOOKUP()
    + Add-ons such as Consolidte Sheets or Merge Sheets 


## The Solution
The solution to this challenge, using VLOOKUP(), was to use a combination of IMPORTRANGE() and VLOOKUP() in a new spreadsheet. IMPORTRANGE() was used import the external Outgoing data (Date, BikeId, Destination) to the BikeCycle spreadsheet. VLOOKUP() was used to search for the corresponding Bicycle attributes in the Incoming spreadsheet using the BikeID as the lookup field.

Let's take a look at how we got around VLOOKUP()'s limitation that requires the lookup field to be in the left most column. If you remember, the BikeId on the Incoming spreadsheet occurs well into the spreadsheet column list and definitely not to the left of the information we want VLOOKUP() to return.

| *Incoming*           |       |                |            |             |
| ------------------ | ----- | -------------- | ---------- | ----------- |
| **Timestamp**      | **Make**  | **Model**      | **Wheel Size** | **Bike ID** |
| 7/29/2022 10:12:14 | Kent  | Ambush         | 20"        |  A0001      |


How do we do this? Using the following formula (it's a long formula and may display on two lines):

```ArrayFormula(IFERROR(VLOOKUP(A2,{Incoming!H2,Incoming!A2:J},{2,3,4,5,6,7,9},False),""))```

*Don't panic!*  
We're going to break this formula into it's individual parts. Let's extract the VLOOKUP() part of this formula and take a look.



```ArrayFormula(IFERROR(``` ***`VLOOKUP(A2,{Incoming!H2,Incoming!A2:J},{2,3,4,5,6,7,9},False)`*** ```,""))```


Let's look at how VLOOKUP() works:

VLOOKUP(lookup_value,table_array,col_index_num,[range_lookup])

You can learn more about VLOOKUP() [here](https://support.microsoft.com/en-us/office/quick-reference-card-vlookup-refresher-750fe2ed-a872-436f-92aa-36c17e53f2ee)