
# XmlToWord

CLI app to generate a Word document from a XML data file.

No need to install Office. :smiley:  
Uses XPath to handle most XML structures. :astonished:  
Built on .NET 4.  :smiling_imp:  


## Example 1

### XML file

    <?xml version="1.0" encoding="utf-8"?>
    <WorkItems>
      <WorkItem Id="80">
        <Fields>
          <Field RefName="System.Description" Value="Setup and start a data collector for the launch and the future." />
          <Field RefName="System.Title" Value="Setup a data collector on the main server" />
        </Fields>
      </WorkItem>
      <WorkItem Id="81">
        <Fields>
          <Field RefName="System.Description" Value="429 Too Many Requests" />
          <Field RefName="System.Title" Value="Develop request threshold on the main website" />
        </Fields>
      </WorkItem>
      ...
    </WorkItems>

### Command line

    XmlToWord.exe ^
      MyDataFile.xml ^
      utf-8 ^
      "C:\Program Files (x86)\Microsoft Office\Office15\1033\QuickStyles\Default.dotx" ^
      GeneratedDoc.docx ^
      /WorkItems/WorkItem ^
      Heading2:./Fields/Field[@RefName='System.Title']:Value: ^
      Paragraph:./Fields/Field[@RefName='System.Description']:Value:

### Remarks

You need to specify the path to a Word template.  
`"C:\Program Files (x86)\Microsoft Office\Office15\1033\QuickStyles\Default.dotx"`

The iteration selector is `"/WorkItems/WorkItem"`. Yes, that's a XPath query.

I build a header of level 2 that will contain the content of the "Value" attribute of the first element which attribute "RefName" is "System.Title".

I build a paragraph that will contain the content of the "Value" attribute of the first element which attribute "RefName" is "System.Description".



## Dependencies

* Install [Open XML SDK 2.5](http://www.microsoft.com/en-us/download/details.aspx?id=30425)

## Possible enhancements

- [ ] Specify the document title in arguments
- [ ] Split into to phases: build variables from XML and then generate document
- [ ] Use a existing Word file and write in a specific position in it. That instead of loading a template.
- [ ] Change ItemStyle from enum to string for custom styles.
- [ ] Add an argument to generate paragraphs instead of line breaks when a newline is found in a value.
- [ ] Allow a recursive process.
- [ ] Interpret a value as markdown and build word content in a smart way.
  - [ ] Generate paragraphs
  - [ ] Generate titles
  - [ ] Generate bullet lists
  - [ ] Generate hyperlinks
- [ ] Allow an argument to filter items based on a value.
- [ ] Make hyperlinks from URLs

### 2 phases process

Here is what it would allow.

    "Set:ID:./Fields/Field[@RefName='System.ID']:Value" ^
    "Set:Date:./Fields/Field[@RefName='System.Date']:Value:System.DateTime" ^
    "Set:Title:./Fields/Field[@RefName='System.Title']:Value" ^
    "Set:Descr:./Fields/Field[@RefName='System.Description']:Value" ^
    "Write:Heading2:{ID} - {Title}" ^
    "Write:Paragraph:Date changed: {Date}" ^
    "Write:Paragraph:{Descr}"

* [ ] Handle the `Set:<id>:<xpath>` argument.
* [ ] Handle the `Set:<id>:<xpath>:<attribute name>` argument.
* [ ] Handle the `Set:<id>:<xpath>:<attribute name>:<type>` argument. This will `Convert.ChangeType` the string value.
* [ ] Handle the `Write:<style>:<format>` argument.
* [ ] For the `Write:<style>:<format>` argument, allow specify a ToString parameter for the object.  
`Date changed: {Date "YYYY-MM-dd"}` or `Number: {ID F2}`

