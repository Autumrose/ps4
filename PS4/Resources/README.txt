Autumrose Stubbs
9/18/2019
Your initial design thoughts about the project: how you are going to set things up/code the project, how are you going to represent cells?
My initial thoughts about how to set up a cell is a pretty basic structure. I'm thinking a separate class in the same namespace, obviously, 
and then a constructor to create the cell and some methods to get and set the contents of the cell. I think this will provide the functionality I 
will need to be able to have them as objects and reference them along with changing them, as of what I know now there isn't anything else I would need to 
be doing in there. For the rest of the class I will have two main structures, one a dictionary to keep track of all of the names plus their cells and 
an instance of dependencies class to keep track of dependencies.

In this class I will be using the most current versions of PS2 and PS3. PS2 has updates all the way up until 9/20/2019 that were changes to fix bugs that 
were found after using the Professor Kopta's grading tests. PS3 also was updated earlier this week for some of the file organization and the double equals
method were incorrect.

My code coverage isn't 100% because I couldn't figure out how the checks for if the text, double, or formula are null could be tested, since the methods 
are overloaded then if you try to call SetCellContents with a string and a null then it doesn't understand which method you are referencing so it won't allow 
it. I still however included these checks just in case. Even if I tried creating a new Formula with null as parameter then it would throw a formula format 
exception and still not enter the check inside this class. 
Also because GetDirectDependents is a private method I couldn't test it directly, and if I had tried to test it indirectly through the method it's used in then
it would've thrown an exception before it ever reached the GetDirectDependents method it's self. Again, I still left the check there just in case.