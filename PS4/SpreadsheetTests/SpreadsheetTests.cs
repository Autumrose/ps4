using SS;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using SpreadsheetUtilities;

namespace SpreadsheetTests
{
    /// <summary>
    /// Set of methods to test the spreadsheet class functions
    /// Author: Autumrose Stubbs
    /// </summary>

    [TestClass()]
    public class SpreadsheetTests
    {
        /// <summary>
        /// Tests for getCellContents and SetContentsOfCell
        /// Use different naming conventions with underscores, letters, and numbers throughout to test
        /// </summary>
        
        [TestMethod()]
        public void getAndSetContentsOfCellString()
        {
            Spreadsheet spreadsheet = new Spreadsheet();
            spreadsheet.SetContentsOfCell("A", "1");
            spreadsheet.SetContentsOfCell("B", "56");
            Assert.AreEqual("1", spreadsheet.GetCellContents("A"));
            Assert.AreNotEqual("1", spreadsheet.GetCellContents("B"));
        }
        [TestMethod()]
        public void getAndSetContentsOfCellNumber()
        {
            Spreadsheet spreadsheet = new Spreadsheet();
            spreadsheet.SetContentsOfCell("A", "3.0");
            spreadsheet.SetContentsOfCell("B", "5.0");
            Assert.AreEqual(3.0, spreadsheet.GetCellContents("A"));
            Assert.AreNotEqual(3.0, spreadsheet.GetCellContents("B"));
        }
        [TestMethod()]
        public void getAndSetContentsOfCellFormula()
        {
            Spreadsheet spreadsheet = new Spreadsheet();
            spreadsheet.SetContentsOfCell("A", "5 + A1");
            spreadsheet.SetContentsOfCell("B", "1");
            Assert.AreEqual(new Formula("5 + A1"), spreadsheet.GetCellContents("A"));
            Assert.AreNotEqual(new Formula("5 + A1"), spreadsheet.GetCellContents("B"));
            
        }
        [TestMethod()]
        public void getAndSetContentsOfCellRepeats()
        {
            Spreadsheet spreadsheet = new Spreadsheet();
            spreadsheet.SetContentsOfCell("A", "A1 + 4");
            spreadsheet.SetContentsOfCell("B", "56");
            spreadsheet.SetContentsOfCell("A", "3");
            Assert.AreEqual("3", spreadsheet.GetCellContents("A"));
            Assert.AreNotEqual("3", spreadsheet.GetCellContents("B"));

        }
        [TestMethod()]
        public void getAndSetContentsOfCellRepeatsFormula()
        {
            Spreadsheet spreadsheet = new Spreadsheet();
            spreadsheet.SetContentsOfCell("A", "5 + A1");
            spreadsheet.SetContentsOfCell("B", "1");
            spreadsheet.SetContentsOfCell("A", "A1");
            Assert.AreEqual(new Formula("A1"), spreadsheet.GetCellContents("A"));
            Assert.AreNotEqual(new Formula("A1"), spreadsheet.GetCellContents("B"));

        }
        [TestMethod()]
        public void getAndSetContentsOfCellRepeatsFormulaDependency()
        {
            Spreadsheet spreadsheet = new Spreadsheet();
            spreadsheet.SetContentsOfCell("A", "5 + A1");
            spreadsheet.SetContentsOfCell("B", "1");
            spreadsheet.SetContentsOfCell("A", "A3 + A5");
            Assert.AreEqual(new Formula("A3 + A5"), spreadsheet.GetCellContents("A"));
            Assert.AreNotEqual(new Formula("A3 + A5"), spreadsheet.GetCellContents("B"));

        }
        [TestMethod()]
        public void getAndSetCellNotInSet()
        {
            Spreadsheet spreadsheet = new Spreadsheet();
            Assert.AreEqual("", spreadsheet.GetCellContents("A"));
        }
        
        /// <summary>
        /// Tests for exceptions
        /// </summary>
        [TestMethod()]
        [ExpectedException(typeof(InvalidNameException))]
        public void getCellContentsNull()
        {
            Spreadsheet spreadsheet = new Spreadsheet();
            spreadsheet.GetCellContents(null);
        }
        [TestMethod()]
        [ExpectedException(typeof(InvalidNameException))]
        public void getCellContentsInvalid()
        {
            Spreadsheet spreadsheet = new Spreadsheet();
            spreadsheet.GetCellContents("!a");
        }
        [TestMethod()]
        [ExpectedException(typeof(InvalidNameException))]
        public void getCellContentsInvalidSecondChar()
        {
            Spreadsheet spreadsheet = new Spreadsheet();
            spreadsheet.GetCellContents("A!");
        }
        [TestMethod()]
        [ExpectedException(typeof(InvalidNameException))]
        public void SetContentsOfCellNameNull()
        {
            Spreadsheet spreadsheet = new Spreadsheet();
            spreadsheet.SetContentsOfCell(null, "a");
        }
        [TestMethod()]
        [ExpectedException(typeof(InvalidNameException))]
        public void SetContentsOfCellNameInvalid()
        {
            Spreadsheet spreadsheet = new Spreadsheet();
            spreadsheet.SetContentsOfCell("1A", "a");
        }
        [TestMethod()]
        [ExpectedException(typeof(CircularException))]
        public void getAndSetContentsOfCellCyclic()
        {
            Spreadsheet spreadsheet = new Spreadsheet();
            spreadsheet.SetContentsOfCell("A", ("5 + A1"));
            spreadsheet.SetContentsOfCell("B", "1");
            spreadsheet.SetContentsOfCell("A1", "A");

        }

        [TestMethod()]
        [ExpectedException(typeof(InvalidNameException))]
        public void getCellContentsInvalidFormula()
        {
            Spreadsheet spreadsheet = new Spreadsheet();
            spreadsheet.SetContentsOfCell("1F", "F");
        }

        /// <summary>
        /// Tests for getNamesOfAllEmptyCells 
        /// </summary>
        [TestMethod()]
        public void getNamesOfAllEmptyCells()
        {
            Spreadsheet spreadsheet = new Spreadsheet();
            spreadsheet.SetContentsOfCell("A", "3.0");
            spreadsheet.SetContentsOfCell("B", "5.0");
            IEnumerator<string> list = spreadsheet.GetNamesOfAllNonemptyCells().GetEnumerator();
            Assert.IsTrue(list.MoveNext());
            Assert.AreEqual(list.Current, "A");
            Assert.IsTrue(list.MoveNext());
            Assert.AreEqual(list.Current, "B");
        }

    }
}