#This report is for reference purposes only as it is customized for a specific purpose and a specific restaurant.# <br>
#Target:<br>
Customer's using TOAST POS and it's Tip report but they are not directly linked together but are two different reports. The goal is to use the data from the two reports to combine into a single report with all the necessary information.<br>
#How to use:<br>
The time-sheet report is placed in the time_src folder and the tip-report will be placed in the tip_src folder. Then, run the main file and the final report will be generated in the output folder.<br>
#time_report<br>
using pandas to filter out all real account and fake account (for general use- not a real person).<br>
Combine users with multiple position but have same payrate.<br>
Adding extra customize payments.<br>
#tips_report<br>
using pandas to filter out all real account and fake account (for general use- not a real person).<br>
Matching user name with their tips<br>
Extract neccessary columns and delete uneccessary informartion.<br>
#combine_report<br>
combine 2 filtered report into a single report.<br>
