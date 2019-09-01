# Genetic Asset Splitter

### This program uses a genetic algorith to find the best way to split assets between divorced couples based on emotional and monetary value of each asset.

### Purpose:

I wrote this program to help out a divorce attorney find ways to split assets between divorced couples. With a problem that has so many solutions to it naturally a genetic algorithm is a great way to find some of the best solutions. The program takes into account an objective monetary value for each asset and a subjective emotional value that each party assigns to each asset. This program is optimized to split the assets 50/50 with the monetary value, make the emotional score for each person the highest and decrease the difference between emotional scores so one person does not have a significant higher score than the other.


### Dependencies

The only dependency is openpyxl to read and create a spreadsheet.
```
pip install openpyxl
```
### How to use:

To use, create a spreasheet with the same format as the sample provided and save it.
Run the python program follow the prompts on the gui and hit start.

When the program is done you will be prompted to save the file with the possible solutions.


How it works

The program uses a genetic algorithm to create possible solutions and then breeds the solutions over many generations until the fitness converges to the highest value.

The fitness function is based on how close the monetary value of the items is split 50-50 between the two people and also how high the emotional value is for each person. However, if the emotional value of one person is higher than the other then the fitness will decrease.
So the fitness function will maximize the emotional score of each person while also keeping them at the same level.

The emotional score of each item is from 1-100 and the program won't use the sum of the emotional score to determine fitness but instead uses a percentage where the emotional score of your alloted items are added and then divided by the sum of the emotional score you put on every item. 

Using the percentage this way ensures the scores are relative to each other and one person will not be shorted items if they put lower scores compared to the other person.


