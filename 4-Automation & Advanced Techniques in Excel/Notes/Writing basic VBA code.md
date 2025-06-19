# Writing basic VBA code                                      
## High-level overview                                                       
While macros are excellent for automating straightforward tasks, Visual Basic for Applications (VBA) provides a more robust solution for advanced automation. Understanding the VBA object model, properties, methods, and syntax is crucial for creating custom solutions tailored to specific needs. This reading introduces these concepts and provides practical insights into writing VBA code, empowering users to harness the full power of Excel.

## Learning objectives                                                                         
By the end of this reading, you will be able to:

explain the VBA object model and its key components

use properties and methods to interact with Excel objects

write basic VBA code to automate tasks effectively

learn how VBA syntax structures instructions for the Excel application

## Understanding the VBA object model           
**What is the object model?**                                      
The VBA object model represents the hierarchy of elements within Excel, from the application itself to specific cells in a worksheet. Each component—referred to as an object—plays a distinct role in the model. Because of their hierarchical structure, objects must always be referenced in the correct order to ensure accurate interactions in VBA.

**Example of object hierarchy**                    
**Application.Workbooks("Sales.xlsx").Worksheets("Sheet1").Range("A1")**                   

In this example:

Application refers to Excel as a whole

Workbooks points to a specific file, e.g., Sales.xlsx

Worksheets identifies a specific sheet within the workbook, e.g., Sheet 1

Range specifies the exact cell or range of cells being referenced, e.g., A1

Note: Objects must be referenced in this exact order—starting from the application and working down through workbooks, worksheets, and ranges. Skipping or misordering any object in the hierarchy will result in errors.

An everyday analogy: Objects, properties, and methods                 
Imagine a car:

Object: The car itself is the object.

Properties: Attributes of the car such as its color, model, or fuel level.

Methods: Actions that the car can perform, like starting the engine, turning on the headlights, or honking the horn.

**In VBA:**

Object: This could be a worksheet or a cell range.

Properties: This might include the value of a cell, the font style, or the column width.

Methods: This includes actions like clearing cell contents, sorting data, or saving the workbook.

This analogy helps clarify how objects, properties, and methods work together in VBA.

## Properties and methods          
**What are properties?**                    
Properties are characteristics or attributes of an object. They allow you to define or retrieve specific details about an object. For example, the Value property of a range object specifies the content of a cell.

**Examples of properties**                     
Value: Sets or retrieves the value of a cell

Range("A1").Value = "Hello, World!"               

Font: Defines font characteristics, such as boldness or size

Range("A1").Font.Bold = True

Interior color: Sets the background color of a cell

Range("A1").Interior.Color = RGB(255, 255, 0) 'Yellow

Column width: Adjusts the width of a column

Columns("A").ColumnWidth = 20

**What are methods?**         
Methods are actions that can be performed on an object. They are used to manipulate objects or execute tasks within Excel.

**Examples of methods**                                          
Clear contents: Removes the content of specified cells

Range("A1").ClearContents

Copy: Copies the content of a range to another location

Range("A1:A5").Copy Destination:=Range("B1")

Sort: Sort a range of data

Range("A1:A10").Sort Key1:=Range("A1"), Order1:=xlAscending

Save as: Saves a workbook with a specified name	

ThisWorkbook.SaveAs "C:\MyWorkbook.xlsx"

**Understanding VBA syntax**          
**What is VBA syntax?**                         
VBA syntax refers to the structure and order of statements in your code. Much like sentence structure in a language such as English, VBA syntax follows a logical order to ensure commands are interpreted correctly by Excel.

For example, in English, a sentence typically requires a noun, verb, and sometimes an object:

Noun: The subject performing the action, e.g., The dog

Verb: The action being performed, e.g., barks

Object: Optional, describing what is affected, e.g., at the cat

**In VBA:**

Object: Specifies what the action applies to, e.g., (Range("A1"))

Property/method: Describes the action or characteristic, e.g., .Value or .ClearContents

Argument: Refers to additional details for the action, e.g., "Hello, World!"

**Example of Syntax in VBA**        
**Sub MyMacro()**                  
   **Range("A1").Value = "Hello, World!"**                  
**End Sub**                  

**This simple macro:**                            

References the object (Range("A1"))

Specifies the property (Value)

Sets the property to a specific value ("Hello, World!")
