#This report is for reference purposes only as it is customized for a specific purpose and a specific restaurant.#
#Target:
Customer's using TOAST POS and it's Tip report but they are not directly linked together but are two different reports. The goal is to use the data from the two reports to combine into a single report with all the necessary information.
#How to use:
The time-sheet report is placed in the time_src folder and the tip-report will be placed in the tip_src folder. Then, run the main file and the final report will be generated in the output folder.
#time_report
using pandas to filter out all real account and fake account (for general use- not a real person).
Combine users with multiple position but have same payrate.
Adding extra customize payments.
#tips_report
using pandas to filter out all real account and fake account (for general use- not a real person).
Matching user name with their tips
Extract neccessary columns and delete uneccessary informartion.
#combine_report
combine 2 filtered report into a single report.