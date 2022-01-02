# oiagj-sheets
Code relating to OIAGJ project spreadsheets

## fixPallet
[``fixPallet.vba``](https://github.com/rkooyenga) is a VBA module to fix the sort on a private client inventory spreadsheet.

#### Challenge

This is a project for a global logistics corporation. One particular location is a warehouse managing inventory for another company on contract and handles shipping raw and finished goods back and forth in addition to storage. A daily generated location inventory spreadsheet contains pallet identifiers as a sort field. When orders are placed based on certain sku and batch number ranges, the specific pallets in those ranges must be selected in a rather exact ascending order given how the inventory is stacked. Given the size of the product and the massive scale of the warehouse if chosen incorrectly large amounts of inventory will needlessly need to be moved out of the way to grab the requested items. This is an easy mistake to make because the naming/numbering scheme of the pallet identifiers do not sort properly. And because those identifiers are used by Fortune 500 companies changing them is impractical.

Up until now the people in charge of organizing daily inventory moves have to take care to go through an out of order list which increases the likelihood of human error which in this case involves forklift drivers doing unnecessary work and hours to move items out of the way or depending on how impractical that is, the pick lists being revised throughout the day to make changes. This can be mostly avoided by taking extra time to eyeball the list for out of sequence numbers, but that's a time consuming pain in the ass when orders need to be filled faster not slower to keep costs down. And again it allows for human error which extends the operating day, and is a challenge to train new employees on.

#### Solution 

Because the pallet identifiers can't be changed system wide, it seemed the easiest solution is to create a new properly sortable additional temporary identifier only for this process and location. That's what this module does.

#### Updates

Still a work in progress as different product ranges have different sort challenges and rules. Once I figure them all out from day to day production use I can clean up the logic. For now though the core issue is solved for +99% of relevant daily tasks with no serious unexpected result or bugs. Added some minor formatting / aesthetic stuff too. 

*Cannot show the daily use examples, input data, screenshots, etc. in any detail for company secret / non-disclosure reasons.*

I'm beyond rusty on Excel. To be honest I've been on Linux so long I haven't used Windows or Microsoft stuff in over a decade so this is my first time writing something in VBA probably since before John Frusciante last quit the Red Hot Chili Peppers 😮

by [**Ray Kooyenga**](https://github.com/rkooyenga)
 
blog and more projects at by [**rkooyenga.github.io**](https://rkooyenga.github.io/)


