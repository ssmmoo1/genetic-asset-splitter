# genetic-asset-splitter
A genetic algorithm that splits assets between divorced couples based on emotional and monetary value

The only dependency is openpyxl
pip install openpyxl

To use, create a spreasheet with the same format as the sample provided and save it.
Run the python program follow the prompts on the gui and hit start.

When the program is done you will be prompted to save the file with the possible solutions.


How it works

The program uses a genetic algorithm to create possible solutions and then breeds the solutions over many generations until the fitness converges to the highest value.

The fitness function is based on how close the monetary value of the items is split 50-50 between the two people and also how high the emotional value is for each person. However, if the emotional value of one person is higher than the other then the fitness will decrease.
So the fitness function will maximize the emotional score of each person while also keeping them at the same level.

The emotional score of each item is from 1-100 and the program won't use the sum of the emotional score to determine fitness but instead uses a percentage where the emotional score of your alloted items are added and then divided by the sum of the emotional score you put on every item. 

Using the percentage this way ensures the scores are relative to each other and one person will not be shorted items if they put lower scores compared to the other person.


