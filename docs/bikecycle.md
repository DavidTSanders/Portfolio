###Purpose
BikeCycle is a Google Sheets based bicycle tracking system for FreeRide, a Do-It-Yourself bicycle cooperative located in Pittsburgh PA.

BikeCycle's purpose is to track the lifecycle of a bicycle through FreeRide from the time a bicycle is first donated, to when it leaves the shop.

###FreeRide
FreeRide accepts donated bicycles in any condition. Volunteers evaluate the bikes, fix them donate them, or take them apart for metal recycling. Once a bike is fixed, it's donated or sold to the public at a drastically reduced price.

## Current Spreadsheets
FreeRide currently uses two speadsheets to record bicycle information, an **Incoming** spreadsheet and an **Outgoing** spreadsheet. 
###Incoming spreadsheet:
 When a bike is donated to the shop a unique number known as a bikeID, is pasted to the frame. The bikeID and other basic information is entered into a form which automatically populates a spreadsheet named Incoming.  An abbreviated example of the Incoming spreadsheet is below:

| *Incoming*           |       |                |            |             |
| ------------------ | ----- | -------------- | ---------- | ----------- |
| **Timestamp**      | **Make**  | **Model**      | **Wheel Size** | **Bike ID** |
| 7/29/2022 10:12:14 | Kent  | Ambush         | 20"        |  A0001      |
| 7/29/2022 10:12:54 | Huffy | Stone Mountain | 20"        |  A0002      |
| 7/29/2022 10:13:32 | Next  | Screamer II    | 20"        |  A0003      |

###Outgoing spreadsheet:
When a bicycle leaves the shop through a donation, sales, or recycling, the bikeId and destination are recorded using a form that populates a spreadsheet named Outgoing.

| *Outgoing*    |                       |
| ------------|---------------------- |
| **Bike ID** | **Outgoing Destination**  |
| A0069       | Volunteer            |
| A0301       | Sold                 |
| A0302       | Communicycle         |



## Proposed Spreadsheet
### BikeCycle spreadsheet:
I was asked to take the two existing spreadsheets, Incoming and Outgoing, and create a third spreadsheet named BikeCycle. The BikeCycle spreadsheet links the Incoming spreadsheet to the Outgoing spreadsheet using the bikeID.


| *BikeCycle* |||||||
| ----------- | -------------------- | ------------------ | ----- | -------------- | ---------- | ----------- |
| **bikeId** | **Outgoing Destination** | **Timestamp**          | **Make**  | **Model**          | **Wheel Size** | **Bike number** |
| A0069       | Volunteer            | 7/29/2022 10:12:14 | Kent  | Ambush         | 20"        |  A0001      |
| A0069       | Volunteer            | 7/29/2022 10:12:54 | Huffy | Stone Mountain | 20"        |  A0002      |
| A0069       | Volunteer            | 7/29/2022 10:13:32 | Next  | Screamer II    | 20"        |  A0003      |

##The Challenges
There were a few challenges to overcome to complete this project. 

+ The two original sheets are in seperate files. 
+ The lookup column in the Incoming spreadsheet, BikeId, is not the first column in the sheet. This presents challenges for the common spreadsheet formula VLOOKUP() which requires that the lookup field be to the left of any fields returned by the search. 

Note: There are other spreadsheet options that can be used, but this is a showcase of my VLOOKUP() skills. Other options include: 


+ INDEX() / MATCH()
+ QUERY()
+ XLOOKUP()
+ Add-ons such as Consolidate Sheets or Merge Sheets 


## The Solution
The solution to this challenge, using VLOOKUP(), is to use a combination of IMPORTRANGE() and VLOOKUP() in a new spreadsheet. IMPORTRANGE() is used to import the Outgoing data (Date, BikeId, Destination) to the BikeCycle spreadsheet. VLOOKUP() is used to search for the corresponding Bicycle attributes in the Incoming spreadsheet using the BikeID as the lookup field.

Let's take a look at how we got around  the VLOOKUP() limitation that requires the lookup field to be in the left most column. If you remember, the bikeId on the Incoming spreadsheet occurs well into the spreadsheet column list and definitely not to the left of the information we want VLOOKUP() to return.

| *Incoming*           |       |                |            |             |
| ------------------ | ----- | -------------- | ---------- | ----------- |
| **Timestamp**      | **Make**  | **Model**      | **Wheel Size** | **Bike ID** |
| 7/29/2022 10:12:14 | Kent  | Ambush         | 20"        |  A0001      |


To avoid the limitations of VLOOKUP() we use the following formula to trick it into using the bikeId.

```ArrayFormula(IFERROR(VLOOKUP(A2,{Incoming!H2,Incoming!A2:J},{2,3,4,5,6,7,9},False),""))```

*Don't panic!*  
We're going to deconstruct this formula into it's individual parts and explain what each part is doing. First, let's highlight the VLOOKUP() part of this formula.

  
```ArrayFormula(IFERROR(``` ***`VLOOKUP(A2,{Incoming!H2,Incoming!A2:J},{2,3,4,5,6,7,9},False)`*** ```,""))```


Next we look at the VLOOKUP() syntax and compare it to the VLOOKUP() in our formula. The Microsoft documentation for Excel defines the syntax for VLOOKUP() as:

```VLOOKUP(lookup_value,table_array,col_index_num,[range_lookup])```

In the syntax above, the *lookup_value* is the search term we are trying to find. The *table_array* is where we expect to find this term and it's related values. The *col_index_num* is the column or columns we want to return. The *[range_lookup]* is optional and specifies, depending on who you ask, whether this is a text or numeric field (numeric is the default, choose False for text).

The table below gives a description of this syntax and shows the values we entered for each part.

| VLOOKUP()| lookup_value | table_array | col_index_num | [range_lookup] | 
| ---------- | ----------- | ------------- | -----------| -------------- |
| Description | What to look for | Where to look | what to return | optional options | 
| Our version |  A2         |  {Incoming!H2,Incoming!A2:J} | {2,3,4,5,6,7,9} | False |


You can learn more about VLOOKUP() [here](https://support.microsoft.com/en-us/office/quick-reference-card-vlookup-refresher-750fe2ed-a872-436f-92aa-36c17e53f2ee)