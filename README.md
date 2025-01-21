# Onward to a Spreadsheet

## Project Overview
This assignment involves building the core internal functionality of a spreadsheet program. The focus is on implementing the "Model" component of the Model-View-Controller (MVC) architecture. This ensures that the internal logic of the spreadsheet is maintainable.
## Design Thoughts
### 1. Internal Logic Separation:
* The core logic (Model) is implemented in a new Spreadsheet class, which extends the AbstractSpreadsheet class.
### 2. Cell Representation:
* Cell class is to encapsulate individual cell properties, including content, value, and dependencies. This abstraction simplifies cell-related operations and facilitates future extensions.
### 3. Dependencies Management:
* The DependencyGraph will be utilized to track relationships between cells, ensuring efficient recalculations when cell contents change.
### 4. Formula Handling:
* The Formula class will be leveraged to represent and evaluate cell formulas.
### Storage for Non-Empty Cells:
* A dictionary, will map cell names to Cell objects, enabling efficient lookup and management of non-empty cells.
### Testing Strategy:
* Unit tests test edge scenarios, invalid inputs, and normal operations. 
