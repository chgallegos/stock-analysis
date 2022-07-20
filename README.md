# VBA stock-analysis
### Refactoring Code Exercise
----
# Overview: Module 2 presented a scenario of creating a VBA code that would take the year input and it would calculate the Total Daily Volumes as well as the yearly Return for every of the 12 stocks that were being analyzed.

The purpose of this challenge is to use a pre existing code that was developed during Module 2, this code was to be refactored as to increase efficiency of the macros in terms of using less memory, thus less processing time.

----

# Results:
##The code

**For Loops** For Loops and Nested For loops were used in order to populate Ticker Data as well as values in the output cells

Insert Image of For Loop

**Conditional Statements** The conditional statements were used in order to allow the loop to keep running based on the appropriate ticker index

Insert Image of Conditional Statement


**Row Count**: It also required for outside research as to find the Row Count (RowCount = Cells(Rows.Count, "A").End(xlUp).Row). 

###Pre-Refactored code:

When the code for Module 2 was ran, before it was refactored it provided the following run times:

**2017**
insert screenshot
**2018**
insert screenshot

###Refactored Code:

Once the code was refactored, the code ran faster, this is because the code was made more efficient by using a shorter code that was more efficient and didnt need as much memory to process a less efficient code.

----

#Summary:

###-What are the advantages or disadvantages of refactoring code?
The advantage that I can see from refactoring a code is that the code becomes more efficient in terms of memory/processing needs, therefore, this could come in very beneficial when doing analysis to enormous amounts of data, in which this usage could influence if wheter there is an actual completion of the job (ie, the computer processor cannot handle the analysis).

A disadvantage that I found was that as I was re-writting the code, it would destabilize it in a way that as I pulled and added lines of code, it seems like it would create gaps and the code would run until I was able to figure out what the new lines of code were doing that were conflicting with the already written ones.

###-How do these pros and cons apply to refactoring the original VBA script?
In this case, these pros and cons were present in refactoring the VBA script, however the pros were not as significan as the processing time, even though it did change substantially, in the big picture the change didnt seem to be of too much relevance.





