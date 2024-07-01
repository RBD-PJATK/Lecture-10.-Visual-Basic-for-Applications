# Lecture 10. Visual Basic for Applications

https://gakko.pjwstk.edu.pl/materialy/2398/lec/wyklad10/index.html

<h3>Abstract</h3>

<p>Lecture 10 is the first one devoted to programming database applications.
We will introduce the fundamental concepts like <i>procedure</i>,
<i>module</i>, <i>class module</i>, <i>event</i> and <i>error handling</i>.
We will show that the objects of the graphical user interface are also
objects in the programming language which have properties and methods. 
We will demonstrate how to use the code editor and the object browser.
We will also explain how to synchronize forms.  Finally, we will list the major
events which usually must be handled by database applications.
</p>

<p><i><a name="VBA">Visual Basic for Applications</a></i> (<i>VBA</i>)
is the programming language intended for the development of  applications of MS Office. 
It is a subset of Visual Basic.  VBA can be used to connect database objects to form
a coherent application.  It contains standard programming concepts like
<i>procedures</i>, <i>variables</i>, statements <i>If</i>, <i>For</i> and <i>Case</i>.</p>

<hr><h3><a name="Typy procedur">Procedures</a></h3>

<p>In MS Access the paradigm of the event-driven programming of database applications 
is realized through procedures.
They are associated with events in the user interface, e.g.:</p>

<ul>
<li>a click of a button placed on a form,
<li>a change of the value of a text box,
<li>opening and closing a form,
<li>printing a report,
<li>opening the database.
</ul>

<p>Every repeatable activity can be automated by means of a procedure. In particular
it concerns actions that can be selected in menus.
For example, you can put a set of buttons onto a form that when clicked
cause the following actions.

<ul>
<li>Display on-line help.
<li>Open another form.
<li>Close the form.
<li>Print the contents of the form.
<li>Move to the next (or previous, first, last,new) record.
<li>Perform some computations and display the result in an unbound field.
<li>Prepare a back-up copy of the database.
<li>Transfer the data to MS Word, MS Excel or another database.
</ul>

<p>There are two kinds of procedures.</p>

<dl>
<dt><b>Function</b>
<dd>It returns a value as its result.  It can be used in expressions. It can also be
	an event procedure in a report or a form.
<dt><b>Sub</b>
<dd>It returns no value.  It can be an event procedure in a report or a form.
</dl>

<h4><a name="example1">Example 1: <i>Square root</i></a></h4>

<p>One of typical programming tasks is to calculate  the square root of a positive number. Here is 
the procedure in VBA that computes this root.</p>

<pre>Function SquareRoot (X As Double) As Double
  Dim Msg As String
  Select Case Sgn(X)	'Compute the sign of the argument
     Case 1             'OK, if positive
       SquareRoot = Sqr(X)
       Exit Function
     Case 0             'Warn the user if zero
       Msg = "0 was passed"
     Case -1	        'Warn the user if negative
       Msg = "Forbidden value"
  End Select
  MsgBox Msg            'Display the error message
End Function</pre>

<p>Here is a task for you.</p>

<table><tr><td class="notec">Write a procedure that solves quadratic
equations with real coefficients.</table> 

<hr><h3><a name="Zmienne">Variables</a></h3>

<p>There are several kinds of variables.</p>

<dl>
<dt><i>Local in a procedure</i>
<dd>They are declared in this procedure keyword <b>Dim</b>.
<dt><i>Local in a module</i>
<dd>They are declared in this module with the keyword <b>Private</b>
	or without this keyword
	(<b>Private</b> is the default visibility specifier).
<dt><i>Global</i>
<dd>They are declared in this module with the keyword <b>Public</b>.  They are accessible 
	everywhere in the application.
</dl>

<hr><h3><a name="Typy danych">Data types of variables</a></h3>

<p>In Visual Basic the names of data types are different from those used elsewhere in MS Access.</p>

<table border="1" align="center">
<tr><th>Access			<th>Visual Basic		<th>Default value
<tr><td><code>Text</code>	<td><code>String</code>		<td><code>&quot;&quot;</code>
<tr><td><code>Number</code>	<td><code>Integer</code>,
				    <code>Long</code>,
				    <code>Double</code>,
				    <code>Single</code>		<td><code>0</code>
<tr><td><code>Currency</code>	<td><code>Currency</code>	<td><code>0</code>
<tr><td><code>Yes/No</code>	<td><code>Boolean</code>	<td><code>False</code>
<tr><td><code>Date/Time</code>	<td><code>Date</code>		<td><code>December 30, 1899</code>
</table>

<p>In VBA there is the special data type <b><code>Variant</code></b>.  It contains the values of all data types.
A variable of this type can be assigned any value.</p>

<p>When you call a procedure, all its arguments are passed by <i>reference</i> by default.
Therefore, inside the procedure we use the same object that has been passed and not its copy.</p>

<p>We can also specify that an argument of a procedure must be passed by value.
To do it, we precede the declaration of this argument by the keyword <b><code>ByVal</code></b>.
When this procedure is called, the passed argument is evaluated and the obtained value
is sent to the procedure.  If this argument is a variable, its value is not modified, even if 
inside the procedure the corresponding argument is altered.</p>

<hr><h3><a name="Modul">Modules</a></h3>

<p><i>A module</i> is a set of declarations and definitions of procedures written in VBA.
It is stored as one whole.</p>

<ul>
<li>A module can contain both event procedure and ordinary named procedures.
<li>There are two kinds of modules.
	<ul>
	<li><i>class modules</i>:
		<ul>
		<li>class modules of database objects like forms and reports,
		<li>class modules that define independent objects.
		</ul>
	<li><i>general modules</i> that are not associated with any database object.
	</ul>
<li>Procedures of a module may be either <code>Public</code> or <code>Private</code>.
	<ul>
	<li><code>Public</code> procedures may be called from anywhere in the application.
	<li><code>Private</code> procedures may be called only from the inside of the module.
	<li>Event procedures are <code>Private</code>.
	</ul>
</ul>

<p>The definition of a module begins with options.  The following options are
added by MS Access automatically when the module is created.</p>

<dl>
<dt><b>Option Compare Database</b>
<dd>The strings will be compared with the method defined in the database.
<dt><b>Option Explicit</b>
<dd>All variables must be declared.
</dl>

<h4><a name="example2">Example 2: <i>Counter</i></a></h4>

<pre>Option Compare Database
Option Explicit

Private counter As Integer    'Local variable
Public  display As Integer    'Global variable

Sub Reset()                   'Global procedure
  counter = 0
  display = counter
End Sub

Sub Increment()               'Global procedure
  counter = counter + 1
  display = counter
  MsgBox counter, , "COUNTER"
End Sub</pre>

<hr><h3><a name="Edytor kodu">Visual Basic Editor</a></h3>

<p><i>Visual Basic Editor</i> offers the environment to run and debug code
written in VBA.  It contains the following tools.</p>

<dl>
<dt><i>Immediate Window</i>
<dd>You can open it by pressing CTRL-G or selecting the menu item
	&quot;View -&gt; Immediate Window&quot;. It allows running code
	(statements, methods, functions and procedures), checking the values of
	expressions, fields and properties. For example, if you write:
	<pre>? counter</pre>
	<p>Visual Basic Editor will display the value of the variable <code>counter</code>.</p>

<dt><i>Breakpoint</i>
<dd>You can place breakpoints in functions and procedures by pressing F9 or
	selecting menu item &quot;Debug -&gt; Toggle Breakpoint&quot;.
<dt><i>Project Explorer</i>
<dd>It allows browsing the classes of the current project.
<dt><i>Object Browser</i>
<dd>It allows browsing registered classes of libraries of MS Access and VBA.

</dl>

<p>Usually one class corresponds to one object only, i.e. <i>the class object</i>.</p>

<hr><h3><a name="db-objects">Accessing database objects</a></h3>

<p>The object that represents the form <i>Employees</i> can be referenced in the following
ways.</p>

<ul>
<li><code><b>Forms</b>![Employees]</code>
<li><code><b>Forms</b>(&quot;Employees&quot;)</code>
<li><code><b>Forms</b>(number)</code><br>
	where <code>number</code> is the sequence number of the form <i>Employees</i> in the current session.
</ul>

<p>We can access properties of this object by means of the dot.</p>

<ul>
<li><code><b>Forms</b>![Employees].Caption</code> is equal to the title of
	the form <i>Employees</i>.
<li><code><b>Forms</b>![Employees].SetFocus</code> is the call that assigns the focus
	to form <i>Employees</i>.
</ul>

<p>The exclamation mark is used to select items from a collection.</p>

<ul>
<li><code><b>Forms</b>![Employees]</code> - select the form <i>Employees</i> from the collection
	of opened forms of the current session.
<li><code><b>Forms</b>![Employees]![Last Name]</code> - select the field <i>Last Name</i>
	from the collection of controls of the form <i>Employees</i>.
</ul>

<p>Therefore,</p>
<ul>
<li>the exclamation selects an item from a collection.
<li>the dot selects a property of an item or a collection.
</ul>

<p>The phrase <code>Forms![Employees]![Last Name]</code> means also the default property of 
this text box, i.e. its value (the given last name). You can also use the unabbreviated 
notation:</p>

<table><tr><td class="notec"><code>Forms![Employees]![Last Name].Value</code></table>

<p>If the user is typing a new value into a text box or a combo box, this value is not yet
stored in the database.  You can access it by means of property <code>Text</code>.</p>

<table><tr><td class="notec"><code>Forms![Employees]![Last Name].Text</code></table>

<p>If you want to use this property, you have to set the focus to the referenced control.</p>

<table><tr><td class="notec"><code>Forms![Employees]![Last Name].SetFocus<br>
MsgBox "Last name = " & Forms![Employees]![Last Name].Text</code></table>

<h4><a name="example3">Example 3: Is the form open?</a><a name="Sprawdzanie"/></h4>

<p>This task can be performed in a number of ways. One of them is to use the standard
function <code>SysCmd</code>.

<pre>
Function IsOpen(ByVal FormName As String) As Integer
  IsOpen = False

  'Is it opened at all? 
  If SysCmd(acSysCmdGetObjectState, acForm, FormName) &lt;&gt; 0 Then

    'Is it opened in the form view?
    If Forms(FormName).CurrentView &lt;&gt; 0 Then 
      IsOpen = True
    End If

  End If
End Function
</pre>


<p>Notice that we distinguish the class of the form with its properties from
the actual instances visible on the screen.  We can perform VBA commands on
the instances only. The instance of the form is called the object of the
form or for short the form. The opened forms are objects of their
classes.</p>

<p>The function which checks whether a form is opened (either in the form
view or in the design view) can be implemented by the scan of the collection
of all opened forms.</p>


<pre>Function IsOpen (ByVal FormName As String) As Integer
' Returns True if form named FormName is opened 
' either in the form view or in  the design view

  Dim I As Integer
  IsOpen = False

  ' Forms.Count is the number of items in collection Forms
  For I = 0 To Forms.Count - 1
    If Forms(I).Name = FormName Then
      IsOpen = True
      Exit For
    End If
  Next I
End Function</pre>

<p>You can create instances of a form in VBA code.</p>

<table><tr><td class="notec"><code>Dim MyCopy As New Employees</code></table>

<p>Then you can perform actions on this form. Initially such an instance is not
visible on the screen.</p>


<hr><h3><a name="Kolejna">Event procedures</a></h3>

<p>Event procedures handle events.  An event procedure is associated with an object
(a form, a control etc.) and an event.  This event procedure is executed,
when this event occurs.</p>

<p>When you fire a wizard, it will create an event procedure.
Let us take a look at the procedure created automatically when 
the developer adds the command
button that opens the form <i>Employees</i>.</p>


<pre>Private Sub Button1_Click()
  On Error GoTo <font color="green"><code>Err_Button1_Click</code></font>
  Dim stDocName As String
  Dim stLinkCriteria As String
  stDocName = "Employees"
  DoCmd.OpenForm stDocName, , , stLinkCriteria

<font color="blue"><code>Exit_Button1_Click</code></font>:
  Exit Sub

<font color="green"><code>Err_Button1_Click</code></font>:
  MsgBox Err.Description
  Resume <font color="blue"><code>Exit_Button1_Click</code></font>
End Sub</pre>


<p>Command <code>DoCmd.OpenForm &quot;Employees&quot;</code> is the call to
the method
<code>OpenForm</code> of the built-in object <code>DoCmd</code>.
This call opens the form
<i>Employees</i>.</p>

<h4>Error handling</h4>

<p>The procedure presented above handles errors. The button wizard creates event procedure
<code>Button1_Click</code> and codes there the following strategy.</p>

<ol>
<li>If an error occurs (i.e. the form <i>Employees</i> cannot be opened), suspend the execution and
	go to the error handling section after the label <font color="green"><code>Err_Button1_Click</code></font>.
<li>Show the message box with the information on this error read by means of
	the method <code>Description</code> of the special built-in object <code>Err</code>.
<li>Stop the execution of the procedure.
</ol>
	

<p>You can put your own error message into this procedure.  Simply, replace <code>Err.Description</code>
with another text for the user, e.g.:</p>

<pre>MsgBox "Application error. Contact the administrator."</pre>

<p>Please solve the following task.</p>

<p align="center">
<table><tr><td class="notec"><a href="javascript:popUp('ok01.html',550,270)">Write</a>
another function that checks whether the form whose name is passed as the argument is opened
(either in the design view of the form view). Use the mechanism of error handling in you solution.
</table> 


<hr><h3><a name="Synchronizacja">Synchronization of forms</a></h3>

<p>Sometimes the opening of a form is more complex and thus cannot be
generated by wizards.  Imagine a button that opens the form <i>Participation in
projects</i> which is to show all the projects of the employee selected on
the subform of the form <i>Departments</i>.  The button is placed on the main
form.</p>

<p align="center"><img border="0" src="https://gakko.pjwstk.edu.pl/materialy/2398/lec/wyklad10/images/10_1.png"></p>

<p>The event procedure of this button opens the form <i>Participation</i>
that shows the list of projects in which the person selected on the subform
is involved.</p>

<pre>Private Sub Projects_Click()
On Error GoTo Err_Projects_Click

  Dim stDocName As String
  Dim stLinkCriteria As String

  stDocName = "Participation"
   
  stLinkCriteria = "[Empno]=" &amp; Forms![Departments]![Person Subform].Form![Empno]
  DoCmd.OpenForm stDocName, , , stLinkCriteria

  Forms![Departments].SetFocus
  Forms![Departments]![Person Subform].SetFocus

Exit_Projects_Click:
  Exit Sub

Err_Projects_Click:
  MsgBox Err.Description
  Resume Exit_Projects_Click
    
End Sub</pre>

<p>The form <i>Participation</i> should show the projects of the employee
who
is currently selected in the subform of the form <i>Departments</i>.  If the
selection is changed, the form <i>Participation</i> must show another set of
rows, i.e. the projects of the newly selected person.  To handle this
situation we have to create the event procedure for the event <i>On Current</i>
in the subform of form <i>Departments</i>.  This event occurs when the user
moves to another row displayed by the form.  The change of the set of rows
displayed by the form <i>Participation</i> makes sense unless this form is
closed. The button <i>Projects</i> may be clicked, if the field <i>Empno</i> of
the subform is empty.</p>

<p>We present appropriate event procedure below.  This procedure uses
two new symbols that are references to objects.</p>

<ul>
<li><b>Me</b> - the form or the report whose event procedure is executed.
<li><b>Parent</b> - the superform or the superreport of 
	the form or the report whose event procedure is executed.
</ul>


<pre>Private Sub Form_Current()
On Error GoTo Err_Form_Current

  If IsNull(Me![Empno]) Then
    Me.Parent![Projects].Enabled = False
  Else
    Me.Parent![Projects].Enabled = True

    If IsOpen("Participation") Then
      Dim stLinkCriteria As String

      stLinkCriteria = "[Empno]=" &amp; Me![Empno]
      DoCmd.OpenForm "Participation", , , stLinkCriteria

      Forms![Departments].SetFocus
      Forms![Departments]![Person Subform].SetFocus
    End If
  End If
   
Exit_Form_Current:
  Exit Sub

Err_Form_Current:
  MsgBox Err.Description
  Resume Exit_Form_Current

End Sub</pre>


<p>The change of the set of rows displayed by the form <i>Participation</i> will
be needed also when the user modifies the value of the field <i>Empno</i>.  This
change may be implemented by the procedure for the event <i>After Update</i> of
the field <i>Empno</i>. Its code will be similar to the code of procedure
<i>Form_Current</i>.</p>


<p>It is time for a task for you.</p>

<table><tr><td class="notec">
<p>Create two tables <i>Employees</i> and <i>Departments</i>.  Connect them with the usual many-one
relationship. Every department employs many employees. Every employee works in exactly
one department.</p>

<p>Create two forms <i>Employees</i> (to display the data on employees)
and <i>Departments</i> (to display the data on departments).</p>

<p>Put the button <i>Employees of department</i> onto the form
<i>Departments</i>. When the user clicks this button, the form <i>Employees</i>
will open and display the employees of departments selected on the form
<i>Departments</i>.  If the user changes the current record in the form
<i>Departments</i> and the form <i>Employees</i> is open, it will automatically
display the employees of the newly selected department.</p>

<p><a href="javascript:popUp('ok02.html',600,390)">Which</a> events should be handled? 
Write appropriate event procedures.
</table> 

<p>And another one.</p>

<table><tr><td class="notec">
<p>Does you solution allow adding new employees by means of the form <i>Employees</i>
when this form is opened through the form <i>Departments</i>?</p>

<p><a href="javascript:popUp('ok03.html',600,230)">Which</a> events should be handled? 
Write appropriate event procedures.
</table> 

<hr><h3><a name="Zdarzenia">Events</a></h3>
 
<p>We will present a number of events that can occur in forms, reports and controls
of MS Access databases.  We have chosen the events that are used most often.</p>

<h4>On Open</h4>

<p><i>Open</i> is an event of a form or a report. It occurs, after the user opens the form or the report
but before the event <a href="#ev-load"><i>Load</i></a>.  Its event procedure can e.g.:</p>

<ul>
<li>close another window,
<li>assign focus to a control,
<li>ask the user which records to display,
<li>ask the user for the password and cancel the opening if the user is unable to type the valid password. 

<pre>Private Sub Form_Open(Cancel As Integer)
  Dim strPass as String
  strPass = InputBox("Enter password:")
  If strPass &lt;&gt; "I love VB" Then
    MsgBox "Invalid password"
    Cancel = True
  End If
End Sub</pre>

</ul>

<h4><a name="ev-load">On Load</a></h4>

<p><i>Load</i> is an event of a form. It occurs, after the just opened form displays the records. 
Its event procedure can e.g.:</p>

<ul>
<li>set the default settings for controls,
<li>compute the fields derived from fields of other forms.
<li>display additional windows.

<pre>Private Sub Form_Load()
  Dim StrName As String
  MsgBox "The form will soon appear."
End Sub</pre>

</ul>


<h4>On Current</h4>

<p><i>Current</i> is an event of a form. It occurs, after a record becomes
the current one when:

<ol>
<li>you navigate to another record,
<li>you open the form (in this case the first record becomes the current one),
<li>you refresh the form.
</ol>

<p>The event procedure for <i>Current</i> can e.g.:</p>

<ul>
<li>display a message on the transition to the new record,
<li>change the properties of controls, e.g. hide or show some of them,
<li>change the form's caption.

<pre>Private Sub Form_Current()
  Forms![Employees].Caption = Me![Last Name]
End Sub</pre>

</ul>

<h4>On Delete</h4>

<p><i>Delete</i> is an event of a form. It occurs, just before a record is physically
deleted. The event procedure for <i>Delete</i> can e.g.:</p>

<ul>
<li>ask for confirmation that the user really wants this record to be deleted.


<pre>Private Sub Form_Delete(Cancel As Integer)
  If MsgBox("Really delete?", vbYesNo) = vbNo Then
    Cancel = True
  End If
End Sub</pre>

</ul>

<h4>On Unload</h4>

<p><i>Unload</i> is an event of a form. It occurs, when the form is about to be closed.
The event procedure for <i>Unload</i> can e.g.:</p>

<ul>
<li>ask for confirmation that the user really wants this from to be closed.
<li>perform additional actions, e.g. write a message into the log.
</ul>

<h4>On Close</h4>

<p><i>Close</i> is an event of a form or a report. It occurs when the form or the report
has already been closed and removed from the screen. The event procedure for
<i>Close</i> can e.g. use the method <code>OpenForm</code> to open another form.</p>

<h4>After Insert</h4>

<p><i>After Insert</i> is an event of a form. It occurs, after a new record
has been added to the database. The event procedure for <i>After Insert</i>
can e.g. refresh data.</p>

<h4>Before Update</h4>

<p><i>Before Update</i> is an event of a form or a control.
It occurs just before an update of a record or a field is performed.
The event procedure for <i>Before Update</i> can e.g. validate
the change and possibly cancel it.</p>

<h4>After Update</h4>

<p><i>After Update</i> is an event of a form or a control.
It occurs, after an update of a record or a field has been performed.
The event procedure for <i>After Update</i> can e.g. apply a filter or refresh
the data.</p>

<h4>On Click</h4>

<p><i>Click</i> is an event of a form or a control.
It occurs, when the user clicks the mouse button over the form or the control.</p> 

<pre>Private Sub cmdClickMe_Click()
  MsgBox "What do you want from me?"
End Sub</pre>

<h4>On Key Press</h4>

<p><i>Key Press</i> is an event of a form or a control.  It occurs, after a
key is pressed. The event procedure for <i>Key Press</i> can e.g. check the
newly entered character or the whole text (it can be read from the property
<code>Text</code>).</p>

<h4>On Got Focus</h4>

<p><i>Got Focus</i> is an event of a form or a control.
It occurs, after the form or the control has got the focus. Properties 
<i>Enabled</i> and <i>Visible</i> of the control which gets focus must be set to
<i>True</i>. 
The event procedure for <i>Got Focus</i> can e.g. set the caption of the text box.</p>

<pre>Private Sub Last_Name_GotFocus()
   [lblLast_Name].Caption = "The customer being considered"
End Sub</pre>

<h4>On Exit</h4>

<p><i>Exit</i> is an event of a form or a control.
It occurs when the form or the control is about to loose the focus.
The event procedure can cancel the exit.
</p>

<h4>On Lost Focus</h4>

<p><i>Lost Focus</i> is an event of a form or a control.
It occurs, after the form or the control has lost the focus.
The event procedure cannot cancel this event but it can, e.g.</p>

<ul>
<li>validate the entered data (here the properties <i>Text</i> and <i>Value</i> are already
	equal),
<li>change the properties of the control.
</ul>

<h4>Switching to another control</h4>

<p>When the user tries to move focus to another control, the following events will occur
(in this sequence).</p>

<ol>
<li><i>Before Update</i> (can be canceled), 
<li><i>After Update</i> (cannot be canceled),
<li><i>Exit</i> (can be canceled),
<li><i>Lost Focus</i> (cannot be canceled).
</ol>

<p>The first two events occur only if the control is dirty, i.e. its content has been changed
since it got focus for the last time.</p>

<p>We can validate the entered data and possibly cancel the event.</p>

<pre>Private Sub txtPass_Exit(Cancel As Integer)
  If Len(txtPass) < 8 Then
    MsgBox "Password must contain at least 8 characters."
    Cancel = True
  End If 
End Sub</pre>

<p>Or we can just notify the user and let him/her go on.</p>

<pre>Private Sub txtPass_LostFocus()
  If Len(txtPass) < 8 Then
    MsgBox "Password must contain at least 8 characters."
  End If 
End Sub</pre>

<p>Now do the following exercise.</p>

<p align="center">
<table><tr><td class="notec">
<p>In this exercise we will test how event procedures work.</p>

<p>Define a table with two columns: <code>id</code> 
of type <code>Autonumber</code> (it is the primary key) and <code>number</code>
of type <code>Integer</code>. Build a form for this table.  Create event procedures
for some events. Start from those presented by this lecture.  The bodies of these
procedures shall contain the calls to <code>MsgBox</code> with the argument
which is the name of the event.</p>

<p>Test this form: open it, enter data, update a row, remove a row, navigate among rows,
close it. In several steps add more and more event procedures and test the form again.</p>

<p>Do the appearing message boxes conform to your expectations?</p>
</table> 

<hr><h3><a name="DLookUp">DLookUp</a></h3>

<p><i>DLookUp</i> is the function that allows retrieving a single value
from the database.</p>

<p>If for example the form <i>Employees</i> requires the name of the department
stored in the table <i>Departments</i>, you can add a derived field with property
<i>Control Source</i> set to:

<pre>=DLookUp("[Name]";"[Departments]";"[Deptno]=Forms![Employees]![Deptno]")</pre>

<p>The same could be achieved if we based that form on the join of the two tables and not
on the single table.</p>

<p>By means of the function <i>DLookUp</i> you can fill a text box with 
a value computed by a select query.  Let us assume that the 
query <i>Gen_Empno</i> is defined
as:

<pre>SELECT IIf(IsNull(Max([Empno])), 1, Max([Empno]) + 1) AS Next_Empno
FROM Employees;</pre>

<p>You can use this query in the source of a control.</p>

<pre>=DLookUp("[Next_Empno]";"[Gen_Empno]")</pre>

<p>Therefore, this method can be used to generate unique values for keys.</p>

<hr><h3><a name="Podsumowanie">Summary</a></h3>


<p>During lecture 10 we introduced the fundamental concepts of the code of a database application.
These are <i>procedure</i>, <i>module</i>, <i>class module</i>, 
<i>event</i> and <i>error handling</i>.
The objects of the graphical user interface are also
objects in the programming language that have properties and methods. 
We showed how to use the Visual Basic Editor and how to write event procedures.</p>

<hr><h3><a name="Slownik">Dictionary</a></h3>

<dl>

<dt><a href="#DLookUp">DLookUp</a>
<dd>The function that allows retrieving a single value from a database and
	executing a query which returns a single value.

<dt><a href="#Zdarzenia">event</a>
<dd>The phenomenon that occurs during the execution of an application.
	The developer can prepare a special handler of the event.
	This handler has the form of an event procedure.

<dt><a href="#Modul">module</a>
<dd>A set of declarations and definitions of procedures written in VBA. 
	It is stored as one whole.  It is an ordinary module or a class module.

<dt><a href="#Typy procedur">procedure</a>
<dd>The fundamental unit of code written in Visual
	Basic for Applications (VBA). It is a <i>subprogram</i>
	or a <i>function</i> (<b>Function</b>).

</dl>

<hr><h3><a name="Zadania">Exercises</a></h3>

<h4>Information</h4>

<p>In order to make the following exercises, you need the database
<i>Library</i> which was created in one of the exercises in lectures 3 and 6.
We will reference the form <i>Books</i>.</p>

<p>Points labeled with stars require more effort.</p>

<h4><a name="Zadanie 1">Exercise 1</a></h4>

<p>We are going to create a form which displays books. By means of this
form the user will browse the books available in the library.  The user
wants also to search books by subjects and authors.  Switch off the control
wizard before you start the development.</p>

<ol>
<li>Create the form <i>Search books</i> based on the table <i>Books</i>  with the
	form wizard (use option <i>AutoForm: columnar</i>).  Switch to the
	design view and place the label <i>Search books by author and
	subject</i> onto the form header.  Add also the command button
	<i>Close form</i>.

<li>Select the item "Build event" from the pop-up menu.  Then select "Code Builder".
	Code the event procedure for the button <i>Close form</i>.  Close the window with
	code and test whether the button works properly.
<li>Add four buttons to the footer. Define their event procedures <i>On click</i>. 
	They will navigate to the first book, the previous book, the next book
	and to the last book. For every button select tab "Format" and call 
	the Picture Builder.  Select icons that are the most suitable from the
	point of view of the end user.  Will a text caption be better than
	a picture?
<li>Add an unbound combo box to the header.  Set its <i>Row Source</i>
	(tab <i>Data</i>) to the table <i>Subjects</i>.  Set <i>Column Count</i>
	(tab <i>Format</i>) to <code>2</code> and 
	<i>Column Widths</i> to <code>0cm;3cm</code>.  Set <i>Name</i>
	(tab <i>Other</i>) to <code>Select subject</code>. 
	When the user selects a subject from combo box <i>Select subject</i>,
	the displayed books will be limited to the books on the selected subject.
	We are going to code it. Switch to tab <i>Event</i> and call the Code Builder
	for the event <i>After Update</i>.  The body of this procedure shall consist of only
	one command:
	
<pre>DoCmd.ApplyFilter "[Subject_Id]=[Forms]![Search books]![Select subject]"</pre>
<li>Add a command button to the header. It will close the current from and open form
	<i>Books</i>. Write the event procedure for the button <i>Close form</i>.  It will
	contain the following actions:
	<dl>
	<dt><code>MsgBox "Message text"</code>
	<dd>Tell the user what is going to happen.
	<dt><code>DoCmd.Close</code>
	<dd>Close the current form.
	<dt><code>DoCmd.OpenForm "Books"</code>
	<dd>Open the form <i>Books</i>.
	</dl>
<li>** Add an unbound combo box <i>Select author</i> to the header.
	When the user selects an author by means of this control, the form will display
	books by this author. First, create the query <i>Authors of books</i> with three columns:
	<i>Author_id</i>, <i>First Name</i> and <i>Last Name</i>.  Persons who are not
	authors must not be shown by this query. Set this query as the <i>Row Source</i>
	of the new combo box. <i>Column Count</i> to 3 and hide the first column of the 
	query. Write the event procedure for <i>After Update</i>.
	The body of this procedure shall consist of only one command:
	
<pre>
DoCmd.ApplyFilter "[ISBN] In (SELECT ISBN FROM Authors" _
                          &amp; " WHERE [Person_id] = Forms![Search books]![Select author])"</pre>

This means that the field <i>ISBN</i> of the form <i>Search books</i> must be a member
	of the set of ISBNs of books authored by the person selected in combo box
	<i>Select author</i>. 
</ol>

<h4><a name="Zadanie 2">Exercise 2</a></h4>

<p>Develop two separate forms <i>Display subjects</i> and <i>Display books</i>
and then synchronize them.  Do not use the form wizard.</p>

<p>The form <i>Display subjects</i> will show up first. It will present the subjects of
all books in the library.  This form will contain the button <i>Display books</i> that
will open the form <i>Display books</i> an place it beside the first from. 
The form <i>Display books</i> will contain books on the selected subject only
(use method <code>DoCmd.FormOpen</code> with the appropriate filter
condition as its argument).</p>

<p>If the current selection of the subject changes, the form <i>Display books</i> shall
change the set of displayed books to the set of books on the newly selected subject
(create event procedure for the event <a href="#FormCurrentProjects"><i>On Current</i></a>;
use the auxiliary function <a href="#example3"><i>IsOpen</i></a>
and the method <code>DoCmd.FormOpen</code>).</p>


<h4><a name="Zadanie 3">Exercise 3</a></h4>

<p>Create a from that will be used by the librarian to log borrowed and returned books.</p>

<h4><a name="Zadanie 4">Exercise 4</a></h4>

<p>Create the form <i>Control panel</i> with buttons which direct users
to all forms developed
so far.  Add the button <i>Exit application</i> that calls the methods 
<code>DoCmd.Quit</code> or <code>Application.Quit</code>.
If some buttons are needed on a form (e.g. to close the form), add them.
</p>
