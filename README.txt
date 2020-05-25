Explanation of the files:

CorpOrgVis.py - the actual code to run, which I generally run through Spyder for troubleshooting as necessary.

CorpOrgVis Base.xlsm - the excel sheet that you input the pairs of the entity (in source -> subsidiary pairs), with the freedom to choose the colours each entity will have. Note I recommend keeping the Placeholder one on the top as it helps keep the top part tidy (it creates an invisible anchor at the top to link the parent entity, a note with the creation date and source file, as well as any other notes such as a legend for the graph at the very top).

CorpOrgVis Base Chart (ver. Demo).graphml - the yEd output file (without modification). In order to have it present nicely: 1. Tools > Fit Node to Label; 2. Layout > One-Click Layout. 
The .jpg is an example of how the graph looks like after the previously mentioned transformations are applied (based on my yEd settings)

I've set up the code so that it outputs this file based on the name of the excel file it draws from, and then the date and time to aid in version control. A copy of that info will be placed within the graph as well.

Uploaded 2020.05.25