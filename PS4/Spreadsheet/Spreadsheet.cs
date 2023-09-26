using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Xml;
using SpreadsheetUtilities;

namespace SS
{
    /// <summary>
    /// An AbstractSpreadsheet object represents the state of a simple spreadsheet.  A 
    /// spreadsheet consists of an infinite number of named cells.
    /// 
    /// A string is a valid cell name if and only if:
    ///   (1) its first character is an underscore or a letter
    ///   (2) its remaining characters (if any) are underscores and/or letters and/or digits
    /// Note that this is the same as the definition of valid variable from the PS3 Formula class.
    /// 
    /// For example, "x", "_", "x2", "y_15", and "___" are all valid cell  names, but
    /// "25", "2x", and "&" are not.  Cell names are case sensitive, so "x" and "X" are
    /// different cell names.
    /// 
    /// A spreadsheet contains a cell corresponding to every possible cell name.  (This
    /// means that a spreadsheet contains an infinite number of cells.)  In addition to 
    /// a name, each cell has a contents and a value.  The distinction is important.
    /// 
    /// The contents of a cell can be (1) a string, (2) a double, or (3) a Formula.  If the
    /// contents is an empty string, we say that the cell is empty.  (By analogy, the contents
    /// of a cell in Excel is what is displayed on the editing line when the cell is selected.)
    /// 
    /// In a new spreadsheet, the contents of every cell is the empty string.
    ///  
    /// The value of a cell can be (1) a string, (2) a double, or (3) a FormulaError.  
    /// (By analogy, the value of an Excel cell is what is displayed in that cell's position
    /// in the grid.)
    /// 
    /// If a cell's contents is a string, its value is that string.
    /// 
    /// If a cell's contents is a double, its value is that double.
    /// 
    /// If a cell's contents is a Formula, its value is either a double or a FormulaError,
    /// as reported by the Evaluate method of the Formula class.  The value of a Formula,
    /// of course, can depend on the values of variables.  The value of a variable is the 
    /// value of the spreadsheet cell it names (if that cell's value is a double) or 
    /// is undefined (otherwise).
    /// 
    /// Spreadsheets are never allowed to contain a combination of Formulas that establish
    /// a circular dependency.  A circular dependency exists when a cell depends on itself.
    /// For example, suppose that A1 contains B1*2, B1 contains C1*2, and C1 contains A1*2.
    /// A1 depends on B1, which depends on C1, which depends on A1.  That's a circular
    /// dependency.
    /// Author: AutumroseStubbs
    /// </summary>
    public class Spreadsheet : AbstractSpreadsheet
    {
        //Create our private variables to keep track of the cells and them
        private Dictionary<string, Cell> cells;
        private DependencyGraph dependencies;
        //Boolean to keep track of if a cell has been modified
        Boolean changed = false;
        public override bool Changed { get { return changed; } protected set { changed = value; } }

        /// <summary>
        /// Constructor
        /// </summary>
        public Spreadsheet() : base(s => true, s => s, "default")
        {
            //Initialize variables a dictionary to track the cells and a dependency instance to keep track of each cell's dependencies
            cells = new Dictionary<string, Cell>();
            dependencies = new DependencyGraph();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="isValid"></param>
        /// <param name="normalize"></param>
        /// <param name="version"></param>
        public Spreadsheet(Func<string, bool> isValid, Func<string, string> normalize, string version) : base(isValid, normalize, version)
        {
            //Initialize variables a dictionary to track the cells and a dependency instance to keep track of each cell's dependencies
            cells = new Dictionary<string, Cell>();
            dependencies = new DependencyGraph();
        }
        /// <summary>
        /// Constructs an abstract spreadsheet by recording its variable validity test,
        /// its normalization method, and its version information.  The variable validity
        /// test is used throughout to determine whether a string that consists of one or
        /// more letters followed by one or more digits is a valid cell name.  The variable
        /// equality test should be used thoughout to determine whether two variables are
        /// equal.
        /// </summary>
        public Spreadsheet(String filePath, Func<string, bool> isValid, Func<string, string> normalize, string version)
            : base(isValid, normalize, version)
        {
            //Initialize variables a dictionary to track the cells and a dependency instance to keep track of each cell's dependencies
            cells = new Dictionary<string, Cell>();
            dependencies = new DependencyGraph();
            try
            {
                //Read in the file
                using (XmlReader file = XmlReader.Create(filePath))
                {
                    //Variables to track the current cell information
                    string name = "";
                    string contents = "";
                    //Iterate through
                    while (file.Read())
                    {
                        if (file.IsStartElement())
                        {
                            switch (file.Name)
                            {
                                //Check what the current element is and act accordingly
                                case "cell":
                                    file.Read();
                                    name = file.ReadContentAsString();
                                    break;

                                case "contents":
                                    file.Read();
                                    contents = file.ReadContentAsString();
                                    SetContentsOfCell(name, contents);
                                    break;
                                //Throw if neither
                                default:
                                    throw new SpreadsheetReadWriteException("Element has to be a cell or its contents!");
                            }
                        }
                    }
                }
            }
            //Catch any exceptions
            catch (Exception e)
            {
                throw new SpreadsheetReadWriteException("Error: " + e.Message);
            }
        }

        /// <summary>
        /// If name is null or invalid, throws an InvalidNameException.
        /// 
        /// Otherwise, returns the contents (as opposed to the value) of the named cell.  The return
        /// value should be either a string, a double, or a Formula.
        /// </summary>
        public override object GetCellContents(string name)
        {
            //Throws exception if name is null or invalid (uses helper method to check)
            if (name == null || !isValidName(name))
            {
                throw new InvalidNameException();
            }
            //If the item isn't in the set, returns an empty string
            if (!cells.ContainsKey(name))
            {
                return "";
            }
            //Finds the cell associated with the given name and returns it
            cells.TryGetValue(name, out Cell cell);
            return cell.getContents();

        }


        /// <summary>
        /// Enumerates the names of all the non-empty cells in the spreadsheet.
        /// </summary>
        public override IEnumerable<string> GetNamesOfAllNonemptyCells()
        {
            //Create a list to keep track of all the cells
            List<String> allCells = new List<String>();
            //Iterate through our global variable that tracks all the cells
            foreach (KeyValuePair<string, Cell> entry in cells)
            {
                //If the cell is not empty adds the cell to the list
                if (!entry.Value.getContents().Equals(""))
                {
                    allCells.Add(entry.Key);
                }
            }
            //Return the full list
            return allCells;
        }
        /// <summary>
        /// If name is null or invalid, throws an InvalidNameException.
        /// 
        /// Otherwise, the contents of the named cell becomes number.  The method returns a
        /// list consisting of name plus the names of all other cells whose value depends, 
        /// directly or indirectly, on the named cell.
        /// 
        /// For example, if name is A1, B1 contains A1*2, and C1 contains B1+A1, the
        /// list {A1, B1, C1} is returned.
        /// </summary>
        protected override IList<string> SetCellContents(string name, double number)
        {
            //Call helper method 
            return SetCells(name, number);
        }
        /// <summary>
        /// If text is null, throws an ArgumentNullException.
        /// 
        /// Otherwise, if name is null or invalid, throws an InvalidNameException.
        /// 
        /// Otherwise, the contents of the named cell becomes text.  The method returns a
        /// list consisting of name plus the names of all other cells whose value depends, 
        /// directly or indirectly, on the named cell.
        /// 
        /// For example, if name is A1, B1 contains A1*2, and C1 contains B1+A1, the
        /// list {A1, B1, C1} is returned.
        /// </summary>
        protected override IList<string> SetCellContents(string name, string text)
        {
            //Return helper method
            return SetCells(name, text);
        }
        /// <summary>
        /// If the formula parameter is null, throws an ArgumentNullException.
        /// 
        /// Otherwise, if name is null or invalid, throws an InvalidNameException.
        /// 
        /// Otherwise, if changing the contents of the named cell to be the formula would cause a 
        /// circular dependency, throws a CircularException, and no change is made to the spreadsheet.
        /// 
        /// Otherwise, the contents of the named cell becomes formula.  The method returns a
        /// list consisting of name plus the names of all other cells whose value depends,
        /// directly or indirectly, on the named cell.
        /// 
        /// For example, if name is A1, B1 contains A1*2, and C1 contains B1+A1, the
        /// list {A1, B1, C1} is returned.
        /// </summary>
        protected override IList<string> SetCellContents(string name, Formula formula)
        {
            //Throws exception if the name is null or invalid
            if (Equals(formula, null))
            {
                throw new ArgumentNullException();
            }
            if (!isValidName(name))
            {
                throw new InvalidNameException();
            }
            if (formula.ToString().Contains("_"))
            {
                throw new FormulaFormatException("Variables can't contain underscores!");
            }
            //Throws circular dependency if circle path
            GetCellsToRecalculate(name);

            //Adds in a new item to the dictionary if it doesn't already contain the name as a key
            if (!cells.ContainsKey(name))
            {
                cells.Add(name, new Cell(name, formula));

            }
            else //Otherwise has to replace the old value with the new value and remove any old dependencies
            {
                cells.TryGetValue(name, out Cell cell);
                Formula oldFormula = new Formula(cell.getContents().ToString());
                foreach (string variables in oldFormula.GetVariables())
                {
                    dependencies.RemoveDependency(name, variables);
                }
                cell.setContents(formula);

            }
            //Add any new dependencies
            foreach (string variable in formula.GetVariables())
            {
                dependencies.AddDependency(name, variable);
            }
            //Checks for circular exception again
            //Create a new list to keep track of the dependents
            List<String> dependents = new List<String>();
            dependents.Add(name);
            //Iterate through the list given from the helper method, add to the new list
            foreach (String temp in GetCellsToRecalculate(name))
            {
                dependents.Add(temp);
            }
            changed = true;
            //Return list of dependents
            return dependents;
        }

        /// <summary>
        /// If name is null, throws an ArgumentNullException.
        /// 
        /// Otherwise, if name isn't a valid cell name, throws an InvalidNameException.
        /// 
        /// Otherwise, returns an enumeration, without duplicates, of the names of all cells whose
        /// values depend directly on the value of the named cell.  In other words, returns
        /// an enumeration, without duplicates, of the names of all cells that contain
        /// formulas containing name.
        /// 
        /// For example, suppose that
        /// A1 contains 3
        /// B1 contains the formula A1 * A1
        /// C1 contains the formula B1 + A1
        /// D1 contains the formula B1 - C1
        /// The direct dependents of A1 are B1 and C1
        /// </summary>
        protected override IEnumerable<string> GetDirectDependents(string name)
        {
            //Throws exception if the name is null or invalid
            if (name == null)
            {
                throw new ArgumentNullException();
            }
            if (!isValidName(name))
            {
                throw new InvalidNameException();
            }
            //Return the dependents of the given name
            return dependencies.GetDependents(name);
        }
        /// <summary>
        /// If name is null, throws an ArgumentNullException.
        /// 
        /// Otherwise, if name isn't a valid cell name, throws an InvalidNameException.
        /// 
        /// Otherwise, returns an enumeration, without duplicates, of the names of all cells whose
        /// values depend directly on the value of the named cell.  In other words, returns
        /// an enumeration, without duplicates, of the names of all cells that contain
        /// formulas containing name.
        /// 
        /// For example, suppose that
        /// A1 contains 3
        /// B1 contains the formula A1 * A1
        /// C1 contains the formula B1 + A1
        /// D1 contains the formula B1 - C1
        /// The direct dependents of A1 are B1 and C1
        /// </summary>
        private IList<string> SetCells(String name, Object contents)
        {
            //Throws exception if the name is null or invalid
            if (contents == null)
            {
                throw new ArgumentNullException();
            }
            if (!isValidName(name))
            {
                throw new InvalidNameException();
            }
            //Throws circular exception 
            GetCellsToRecalculate(name);

            //Adds in a new item to the dictionary if it doesn't already contain the name as a key
            if (!cells.ContainsKey(name))
            {
                cells.Add(name, new Cell(name, contents));
            }
            else //Otherwise has to replace the old value with the new value and remove any old dependencies
            {
                cells.TryGetValue(name, out Cell cell);
                string old = cell.getContents().ToString();
                Formula f = new Formula(old);
                foreach (string variables in f.GetVariables())
                {
                    dependencies.RemoveDependency(name, variables);
                }
                cell.setContents(contents);
            }
            //Checks for circular exception again
            //Create a new list to keep track of the dependents
            List<String> dependents = new List<String>();
            dependents.Add(name);
            //Iterate through the list given from the helper method, add to the new list
            foreach (String temp in GetCellsToRecalculate(name))
            {
                dependents.Add(temp);
            }
            changed = true;
            return dependents;
        }
        /// <summary>
        /// Helper method to check if a name is valid
        /// This means that it starts with a letter or underscore and is followed only by those or a number
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        private bool isValidName(string name)
        {
            //Throws exception if name is null
            if (name == null)
            {
                throw new InvalidNameException();
            }
            Char[] charArray = new Char[10];
            //Turn the current token into a char array to check for proper variable structure
            charArray = name.ToCharArray();
            String letter = charArray[0].ToString();
            int index = 0;
            for (int i = 0; i < charArray.Length; i++)
            {
                letter = charArray[i].ToString();
                if (!Regex.IsMatch(letter, @"^[a-zA-Z]+$"))
                {
                    break;
                }
                index++;
            }
            for (int i = index; i < charArray.Length; i++)
            {
                letter = charArray[i].ToString();
                if (!Double.TryParse(letter, out Double value))
                {
                    return false;
                }
            }
            return true;
        }
        /// <summary>
        /// Returns the version information of the spreadsheet saved in the named file.
        /// If there are any problems opening, reading, or closing the file, the method
        /// should throw a SpreadsheetReadWriteException with an explanatory message.
        /// </summary>
        public override string GetSavedVersion(string filename)
        {
            //Create a string to return
            String savedVersion = "";
            try
            {
                //Read in the file
                using (XmlReader file = XmlReader.Create(filename))
                {
                    //Parse through the file
                    while (file.Read())
                    {
                        //Check for when the element equals spreadsheet
                        if (file.IsStartElement() && file.Name.Equals("spreadsheet"))
                        {
                            //Assign the string and break
                            savedVersion = file.GetAttribute("version");
                            break;
                        }
                    }
                }
            }
            //Catch any exceptions and return the error message
            catch (Exception e)
            {
                throw new SpreadsheetReadWriteException(e.Message);
            }
            //Return the assigned string
            return savedVersion;
        }

        /// <summary>
        /// Writes the contents of this spreadsheet to the named file using an XML format.
        /// The XML elements should be structured as follows:
        /// 
        /// <spreadsheet version="version information goes here">
        /// 
        /// <cell>
        /// <name>
        /// cell name goes here
        /// </name>
        /// <contents>
        /// cell contents goes here
        /// </contents>    
        /// </cell>
        /// 
        /// </spreadsheet>
        /// 
        /// There should be one cell element for each non-empty cell in the spreadsheet.  
        /// If the cell contains a string, it should be written as the contents.  
        /// If the cell contains a double d, d.ToString() should be written as the contents.  
        /// If the cell contains a Formula f, f.ToString() with "=" prepended should be written as the contents.
        /// 
        /// If there are any problems opening, writing, or closing the file, the method should throw a
        /// SpreadsheetReadWriteException with an explanatory message.
        /// </summary>
        public override void Save(string filename)
        {
            String contents = "";
            try
            {
                XmlWriterSettings settings = new XmlWriterSettings();
                settings.Indent = true;
                settings.NewLineOnAttributes = true;

                using (XmlWriter writer = XmlWriter.Create(filename, settings))
                {
                    //Iterate through all the cells in the spreadsheet
                    foreach (KeyValuePair<string, Cell> cell in cells)
                    {
                        writer.WriteStartDocument();
                        writer.WriteStartElement("spreadsheet");
                        writer.WriteAttributeString("version", Version);
                        //Figure out what content's type is and accordingly convert to string
                        if (Double.TryParse(cell.Value.getContents().ToString(), out Double value))
                        {
                            contents = value.ToString();
                        }else if (cell.Value.getContents() is Formula)
                        {
                            contents = cell.Value.getContents().ToString();
                        }else 
                        {
                            contents = cell.Value.getContents().ToString();
                        }
                        //Write titles and elements after successfully converting and end
                        writer.WriteStartElement("cell");
                        writer.WriteElementString("name", cell.Key);
                        writer.WriteElementString("contents", contents);
                        writer.WriteEndElement();
                        writer.WriteEndDocument();
                    }
                    writer.Close();
                }
            }
            //Catch any exceptions and throw
            catch (Exception e)
            {
                throw new SpreadsheetReadWriteException(e.Message);
            }
        }
        /// <summary>
        /// If name is null or invalid, throws an InvalidNameException.
        /// 
        /// Otherwise, returns the value (as opposed to the contents) of the named cell.  The return
        /// value should be either a string, a double, or a SpreadsheetUtilities.FormulaError.
        /// </summary>
        public override object GetCellValue(string name)
        {
            //Throw exception if cell isn't valid
            if(!isValidName(name) || name == null)
            {
                throw new InvalidNameException();
            }
            //Find the cell associated with the name
            cells.TryGetValue(name, out Cell cell);
            //Create a new formula
            Formula formula = new Formula(cell.getContents().ToString());
            return formula.Evaluate(Lookup);
        }
        /// <summary>
        /// If content is null, throws an ArgumentNullException.
        /// 
        /// Otherwise, if name is null or invalid, throws an InvalidNameException.
        /// 
        /// Otherwise, if content parses as a double, the contents of the named
        /// cell becomes that double.
        /// 
        /// Otherwise, if content begins with the character '=', an attempt is made
        /// to parse the remainder of content into a Formula f using the Formula
        /// constructor.  There are then three possibilities:
        /// 
        ///   (1) If the remainder of content cannot be parsed into a Formula, a 
        ///       SpreadsheetUtilities.FormulaFormatException is thrown.
        ///       
        ///   (2) Otherwise, if changing the contents of the named cell to be f
        ///       would cause a circular dependency, a CircularException is thrown,
        ///       and no change is made to the spreadsheet.
        ///       
        ///   (3) Otherwise, the contents of the named cell becomes f.
        /// 
        /// Otherwise, the contents of the named cell becomes content.
        /// 
        /// If an exception is not thrown, the method returns a list consisting of
        /// name plus the names of all other cells whose value depends, directly
        /// or indirectly, on the named cell. The order of the list should be any
        /// order such that if cells are re-evaluated in that order, their dependencies 
        /// are satisfied by the time they are evaluated.
        /// 
        /// For example, if name is A1, B1 contains A1*2, and C1 contains B1+A1, the
        /// list {A1, B1, C1} is returned.
        /// </summary>
        public override IList<string> SetContentsOfCell(string name, string content)
        {
            //Throw exception if contents equal null or the name is invalid
            if (content == null)
            {
                throw new ArgumentNullException("Cell content can't be null");
            }
            if (name == null || !isValidName(name))
            {
                throw new InvalidNameException();
            }
            //Checks if the element is a double and adds if so
            if (Double.TryParse(content, out Double value))
            {
                return SetCellContents(Normalize(name), value);
            }
            //Otherwise checks if the element starts with an equals
            content = Normalize(content.Trim());
            if (content.StartsWith("="))
            {
                //Convert the element into a formula
                return SetCellContents(Normalize(name), new Formula(content.Substring(1), Normalize, IsValid));
            }
            //Otherwise you can just set the content normally
            else
            {
                return SetCellContents(Normalize(name), content);
            }
        }
        public double Lookup(String name)
        {
            cells.TryGetValue(name, out Cell cell);
            Double.TryParse(cell.getContents().ToString(), out Double value);
            return value;
        }
        /// <summary>
        /// Private class to keep track of a cell's name and contents
        /// </summary>
        private class Cell
        {
            String name;
            Object contents;
            /// <summary>
            /// Constructor for cell with double content
            /// Value of this cell is the same as the content
            /// </summary>
            /// <param name="_name"></param>
            /// <param name="_content"></param>
            public Cell(string name, object contents)
            {
                this.name = name;
                this.contents = contents;
            }
            //Set new value to the cell
            public void setContents(Object content)
            {
                this.contents = content;
            }
            //Get the content of the cell
            public Object getContents()
            {
                return contents;
            }
            
        }
    }

}