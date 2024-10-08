# `stdVBA` Examples

This repository holds examples of using `stdVBA`. This should give people a better idea of how to use `stdVBA` and libraries.

## Contents

<table>
  <tr>
    <th>Title</th>
    <th>Type</th>
    <th>Short Description</th>
    <th>Tags</th>
    <th>Dependencies</th>
    <th>Status</th>
  </tr>
  <tr>
    <td><a href="Examples/Inspector-Accessibility-v1">Accessibility Inspector</a></td>
    <td>DevTool</td>
    <td>Inspect the accessibility information of controls UI controls. Critical for UI automation.</td>
    <td>ui, window, automation, embedding,simple</td>
    <td>stdAcc, stdProcess, stdWindow, stdICallable</td>
    <td>Deprecated</td>
  </tr>
  <tr>
    <td><a href="Examples/Inspector-Accessibility-v2">Accessibility Inspector v2</a></td>
    <td>DevTool</td>
    <td>Inspect the accessibility tree of windows. Critical for UI automation.</td>
    <td>ui, window, automation, embedding,advanced</td>
    <td>stdAcc,stdCallback,stdClipboard,stdICallable,stdImage,stdLambda,stdProcess,stdShell,stdWindow</td>
    <td>Complete/WIP</td>
  </tr>
  <tr>
    <td><a href="Examples/Inspector-JSON">JSON Viewer</a></td> 
    <td>DevTool</td>
    <td>Inspect JSON data in a userform.</td>
    <td>json, viewer, browser,advanced</td>
    <td>stdCallback, stdICallable, stdJSON, stdLambda, tvTree</td>
    <td>Complete</td>
  </tr>
  <tr>
    <td><a href="Examples/Inspector-Registry">A registry viewer for VBA</a></td>
    <td>DevTool</td>
    <td>Inspect information in the registry without the need for regedit access.</td>
    <td>registry, win32, viewer, browser,advanced</td>
    <td>stdClipboard, stdIcallable, stdLambda, stdReg</td>
    <td>Complete</td>
  </tr> 
  <tr>
    <td><a href="Examples/Inspector-Clipboard">Clipboard Inspector</a></td>
    <td>DevTool</td>
    <td>Inspect information in the system-wide clipboard.</td>
    <td>clipboard, text, image, binary</td>
    <td>stdClipboard</td>
    <td>Complete</td>
  </tr>
  <tr>
    <td><a href="Examples/Inspector-RunningObjectTable">Running Object Table Inspector</a></td>
    <td>DevTool</td>
    <td>Inspect information in the system-wide Running Object Table.</td>
    <td>component, object, model, COM, running, object, table, ROT</td>
    <td>stdCOM</td>
    <td>Complete</td>
  </tr>
  <tr>
    <td><a href="Examples/BrowserAutomation">Automate the web with Chrome</a></td>
    <td>Library</td>
    <td>Automate major browsers for web scraping or UI presentation.</td>
    <td>web, automation, accessibility, library</td>
    <td>stdAcc, stdEnumerator, stdLambda, stdProcess, stdWindow, stdICallable</td>
    <td>Complete</td>
  </tr>
  <tr>
    <td><a href="Examples/Document Generator">Create documents based on tabular data</a></td> 
    <td>Data Utility</td>
    <td>Build a userform dynamically to make changes to an input object.</td>
    <td>automation, data, document, generator, library, table,advanced</td>
    <td>stdCallback, stdCOM, stdEnumerator, stdICallable, stdLambda, stdRegex, stdTable, stdWindow</td>
    <td>Complete</td>
  </tr>
  <tr>
    <td><a href="Examples/DynamicForm-TransformObject">Dynamically Userform Example</a></td>
    <td>Library</td>
    <td>Build a userform dynamically to make changes to an input object.</td>
    <td>User,Interface,UI</td>
    <td>stdUIElement, stdCallback, stdCOM, stdICallable, stdLambda</td>
    <td>Complete</td>
  </tr>
  <tr>
    <td><a href="Examples/MacroDispatcher">Macro Dispatcher</a></td>
    <td>DevTool / AdminTool</td>
    <td>Define exactly which macros should be run in which particular order across multiple spreadsheets. Monitors success and errors.</td>
    <td>macro, scheduler, dispatcher, bulk, automation, data, process,advanced</td>
    <td>stdAcc, stdCallback, stdEnumerator, stdICallable, stdLambda, stdPerformance, stdReg, stdWindow</td>
    <td>Complete</td>
  </tr> 
  <tr>
    <td><a href="Examples/Notepad-GetAllTextAndClose">Notepad GetAllTextAndClose Utility</a></td>
    <td>Utility</td>
    <td>Get all text from all open Notepad tabs and close them. Useful when you use Notepad for note taking.</td>
    <td>notepad, utility, automation, data, extract</td>
    <td>stdAcc, stdICallable, stdLambda, stdProcess, stdWindow</td>
    <td>Complete</td>
  </tr>
  <tr>
    <td><a href="Examples/SplitSideBySide">Split windows side by side</a></td>
    <td>Productivity Utility</td>
    <td>Split your excel window in half for increased productivity.</td>
    <td>window, automation</td>
    <td>stdLambda, stdWindow, stdICallable</td>
    <td>Complete</td>
  </tr> 
  <tr>
    <td><a href="Examples/Spreadsheet Extractor">Bulk Spreadsheet Extractor</a></td>
    <td>Data Utility</td>
    <td>Extract information from multiple structured Excel documents to a table. Includes version differentiation.</td>
    <td>data, extract, bulk</td>
    <td>stdArray, stdCallback, stdCOM, stdEnumerator,stdICallable, stdLambda, stdPicture, stdRegex</td>
    <td>Complete</td>
  </tr> 
  <tr>
    <td><a href="Examples/stdTable">Further SQL-like table operations and class for handling all things table powered by stdEnumerator</a></td>
    <td>Library</td>
    <td>A sql-like system for handling Excel ListObjects.</td>
    <td>sql, listobject, table, update, select, database</td>
    <td>stdICallable,stdEnumerator,stdCallback,stdJSON</td>
    <td>Complete</td>
  </tr>
  <tr>
    <td><a href="Examples/Timer">Timer</a></td>
    <td>Library</td>
    <td>Call code at set intervals.</td>
    <td>timer, time, stopwatch, watch, poll</td>
    <td>stdCallback, stdICallable</td>
    <td>Complete</td>
  </tr> 
  <tr>
    <td><a href="Examples/uiTextBoxEx-WordControl">Embed Microsoft Word into a userform</a></td>
    <td>UI Control</td>
    <td>Present text to users as a Microsoft Word control.</td>
    <td>ui, window, automation, embedding</td>
    <td>stdLambda, stdWindow, stdICallable, stdProcess</td>
    <td>Complete</td>
  </tr> 
  <tr>
    <td><a href="Examples/xlVBA/xlApplication">Excel Application helper</a></td>
    <td>Library</td>
    <td>Make getting the Application object easier, and extend it with additional methods.</td>
    <td>application, ebmode, edit, mode, LPenHelper</td>
    <td>stdAcc, stdICallable</td>
    <td>WIP</td>
  </tr>
  <tr>
    <td><a href="Examples/xlVBA/xlSaveHandler">Excel custom save handler</a></td>
    <td>Library</td>
    <td>Custom Excel save handler with more events.</td>
    <td>save, handler, custom, events</td>
    <td>None</td>
    <td>Complete</td>
  </tr>
  <tr>
    <td><a href="Examples/xlVBA/xlTableTools">Excel sql-like table operations</a></td>
    <td>Library</td>
    <td>A sql-like system for handling Excel ListObjects.</td>
    <td>sql, listobject, table, update, select, database</td>
    <td>stdICallable</td>
    <td>Complete</td>
  </tr>

  <!-- 
  <tr>
    <td><a href="Examples/">xxx</a></td>
    <td>xxx</td>
    <td>xxx</td>
  </tr> 
  -->
</table>
