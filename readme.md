<h1>Basic VBS (Visual Basic Script)</h1>
<h3>How to run code?</h3>
<ol>
  <li><h5>Copy and paste code in NotePad</h5></li>
  <li><h5>Save file name ->  filename.vbs</h5></li>
  <li><h5>Double Cilck file .vbs for run</h5></li>
</ol>
<Br>

Lesson : 
  <ol>
    <li><a href = "#Variable">Variable</a></li>
    <li><a href = "#Operator">Operator</a></li>
    <li><a href = "#Array">Array</a></li>
    <li><a href = "#If">If-else</a></li>
    <li><a href = "#Loop">Loop</a></li>
    <li><a href = "#Function">Function</a></li>
  </ol>
 
### Print Output :
```bash
Dim text
text = "Hello, World!"

' Get the length of the string'
WScript.Echo "The length of the text is: " & Len(text)

wscript.echo "Hello world" & 10
```
<h6>notice : </h6>
<h6>1. &  is connecter </h6>
<h6>2. Use the Len() function to find the length of a string. </h6>

### Comment
```bash
' this is comment.
```
<br>


<h1 id ="Variable"> Variable </h1>

<i>Syntax</i>

```bash
dim Variable1, Variable2,...  
set Variable1, Variable2,...
```      
<b>Dim</b> <h6> Variable of dim collect data type ->  <b>number , character , string , array</b> </h6>
<b>Set</b> <h6> Variable of set collect data type ->  <b>Object</b> </h6>

<i>Example dim : 

```bash
  dim  mynumber , mystring 
      mynumber = 10
      mystring = "hello"
  wscript.echo mynumber
  wscript.echo mystring
```      
Example set : 

```bash
  set obj = CreateObject("wscript.shell")
  obj.run "cmd.exe"
  set obj = nothing     'clean up for release memory
```
</i>
<br>
<h2 id="Operator">Operator</h2>

 <table>
        <thead>
            <tr>
                <th>Operator</th>
                <th>Precedence (Highest First)</th>
                <th>Description</th>
                <th>Example</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td>^</td>
                <td>1</td>
                <td>Exponentiation</td>
                <td><code>result = 2 ^ 3 ' result is 8</code></td>
            </tr>
            <tr>
                <td>-</td>
                <td>2</td>
                <td>Unary Negation</td>
                <td><code>result = -5 ' result is -5</code></td>
            </tr>
            <tr>
                <td>*</td>
                <td>3</td>
                <td>Multiplication</td>
                <td><code>result = 4 * 2 ' result is 8</code></td>
            </tr>
            <tr>
                <td>/</td>
                <td>3</td>
                <td>Division</td>
                <td><code>result = 10 / 2 ' result is 5</code></td>
            </tr>
            <tr>
                <td>\</td>
                <td>3</td>
                <td>Integer Division</td>
                <td><code>result = 10 \ 3 ' result is 3</code></td>
            </tr>
            <tr>
                <td>Mod</td>
                <td>3</td>
                <td>Modulus (Remainder)</td>
                <td><code>result = 10 Mod 3 ' result is 1</code></td>
            </tr>
            <tr>
                <td>+</td>
                <td>4</td>
                <td>Addition</td>
                <td><code>result = 5 + 3 ' result is 8</code></td>
            </tr>
            <tr>
                <td>-</td>
                <td>4</td>
                <td>Subtraction</td>
                <td><code>result = 5 - 3 ' result is 2</code></td>
            </tr>
        </tbody>
    </table>

<i>Example : 

```bash 
  dim ten , five
  ten = 20-10
  five = 3+7
  five = 5
  wscript.echo ten & ", " & five
```
</i><br>    

<h1 id ="Array"> Array </h1>

### 1. Declare Array 
<h6> Fixed-Size Arrays
Use the Dim statement to declare arrays with a fixed size.</h6>

```bash
Dim arr(4) ' Declares an array with 5 elements (0 to 4)
```
<h6> Dynamic Arrays
Use Dim to declare, and then use ReDim to resize later.</h6>

```bash
Dim arr() ' Declares a dynamic array'
ReDim arr(5) ' Resizes the array to hold 6 elements (0 to 5)

```
### 2. Assigning Valuses 
<h6>Assigning Individual Elements</h6>

```bash
Dim arr(2)
arr(0) = "Apple"
arr(1) = "Banana"
arr(2) = "Cherry"

```
<h6>Assigning Multiple Values Using Array</h6>

```bash
Dim arr
arr = Array("Apple", "Banana", "Cherry")
```
<h4>2.1  Release memory</h4>

```bash
' Clear the array using Erase'
Erase arr
```
<h4>2.2 Size of Array</h4>

```bash
Dim arr
arr = Array(10, 20, 30, 40)

' Calculate the size of the array'
Dim size
size = UBound(arr) - LBound(arr) + 1

WScript.Echo "The size of the array is: " & size

```
### 3.  Accessing Array Elements
```bash
WScript.Echo arr(0) ' Outputs: Apple'
WScript.Echo arr(1) ' Outputs: Banana
```

### 4. Dynamic Array with ReDim Preserve
```bash
Dim arr()
ReDim arr(2)
arr(0) = "Apple"
arr(1) = "Banana"
arr(2) = "Cherry"

ReDim Preserve arr(4) ' Expands the array to 5 elements while keeping values'
arr(3) = "Date"
arr(4) = "Elderberry"
```
### 5. Looping Through an Array
<h6>Using for...next</h6>

```bash
Dim arr
arr = Array("Apple", "Banana", "Cherry")

Dim i
For i = 0 To UBound(arr)
    WScript.Echo arr(i)
Next

```
<h6>Using for..Each</h6>

```bash
Dim arr
arr = Array("Apple", "Banana", "Cherry")

Dim fruit
For Each fruit In arr
    WScript.Echo fruit
Next

```
### 6. Array Functions
 <table>
        <thead>
            <tr>
                <th>Function</th>
                <th>Description</th>
                <th>Example</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td><code>UBound(array)</code></td>
                <td>Returns the upper bound (highest index).</td>
                <td><code>UBound(arr)</code> = <code>2</code> (for 3 elements)</td>
            </tr>
            <tr>
                <td><code>LBound(array)</code></td>
                <td>Returns the lower bound (always 0 in VBScript).</td>
                <td><code>LBound(arr)</code> = <code>0</code></td>
            </tr>
            <tr>
                <td><code>IsArray(variable)</code></td>
                <td>Checks if a variable is an array.</td>
                <td><code>IsArray(arr)</code> = <code>True</code></td>
            </tr>
        </tbody>
    </table>

### 7. Multidimensional Arrays
<h6>Fixed-Size Multidimensional Array</h6>

```bash
Dim arr(2, 2) ' 2D array (3x3 matrix: indices 0-2 for each dimension)'
arr(0, 0) = "Apple"
arr(0, 1) = "Banana"
arr(1, 0) = "Cherry"
arr(1, 1) = "Date"
```
<h6>Accessing Multidimensional Arrays</h6>

```bash
WScript.Echo arr(0, 0) ' Outputs: Apple'
WScript.Echo arr(1, 1) ' Outputs: Date

```
<h6>Dynamic Multidimensional Array</h6>

```bash
Dim arr()
ReDim arr(2, 2) ' Resize to a 3x3 array

```
<br> <hr>   
<i>Example Full Array Usage: 

```bash 
Dim fruits()
ReDim fruits(2)

fruits(0) = "Apple"
fruits(1) = "Banana"
fruits(2) = "Cherry"

' Add more fruits while keeping old values'
ReDim Preserve fruits(4)
fruits(3) = "Date"
fruits(4) = "Elderberry"

Dim i
For i = 0 To UBound(fruits)
    WScript.Echo "Fruit " & i & ": " & fruits(i)
Next

```
Output: 

```bash 
Fruit 0: Apple
Fruit 1: Banana
Fruit 2: Cherry
Fruit 3: Date
Fruit 4: Elderberry
```
</i> 
<br>

<h1 id="If">IF-Else Statment</h1>

  <table>
        <thead>
            <tr>
                <th>Operator</th>
                <th>Description</th>
                <th>Example</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td>=</td>
                <td>Equal to</td>
                <td>If a = b Then</td>
            </tr>
            <tr>
                <td>&lt;&gt;</td>
                <td>Not equal to</td>
                <td>If a &lt;&gt; b Then</td>
            </tr>
            <tr>
                <td>&lt;</td>
                <td>Less than</td>
                <td>If a &lt; b Then</td>
            </tr>
            <tr>
                <td>&gt;</td>
                <td>Greater than</td>
                <td>If a &gt; b Then</td>
            </tr>
            <tr>
                <td>&lt;=</td>
                <td>Less than or equal to</td>
                <td>If a &lt;= b Then</td>
            </tr>
            <tr>
                <td>&gt;=</td>
                <td>Greater than or equal to</td>
                <td>If a &gt;= b Then</td>
            </tr>
        </tbody>
    </table>
      <table>
        <thead>
            <tr>
                <th>Operator</th>
                <th>Description</th>
                <th>Example</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td>And</td>
                <td>True if <strong>both</strong> conditions are True</td>
                <td><code>If a > 0 And b > 0 Then</code></td>
            </tr>
            <tr>
                <td>Or</td>
                <td>True if <strong>either</strong> condition is True</td>
                <td><code>If a > 0 Or b > 0 Then</code></td>
            </tr>
            <tr>
                <td>Not</td>
                <td>Reverses the logical value (True to False)</td>
                <td><code>If Not a > 0 Then</code></td>
            </tr>
        </tbody>
    </table><br>

    
<i>Syntax : 

```bash 
  If condition Then
    ' Code block for True condition'
ElseIf another_condition Then
    ' Code block for another True condition'
Else
    ' Code block for False condition'
End If

```
</i><br>    
<i>Example If-Else : 

```bash
Dim num
num = 0

If num > 0 Then
    WScript.Echo "Positive number."
Else
    WScript.Echo "Number is zero."
End If

```
</i><br>    
</i> 
<i>Example If- Else If - Else : 

```bash
Dim num
num = 0

If num > 0 Then
    WScript.Echo "Positive number."
ElseIf num < 0 Then
    WScript.Echo "Negative number."
Else
    WScript.Echo "Number is zero."
End If

```
</i><br>    

<h1 id ="Loop"> Loop Statment</h1>

<table>
        <thead>
            <tr>
                <th>Loop Type</th>
                <th>When to Use</th>
                <th>Key Feature</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td>For...Next</td>
                <td>When the number of iterations is known in advance</td>
                <td>Can use <code>Step</code> for custom increments/decrements.</td>
            </tr>
            <tr>
                <td>For Each...Next</td>
                <td>To iterate through arrays or collections</td>
                <td>Simplifies iterating through groups of items.</td>
            </tr>
            <tr>
                <td>Do...Loop</td>
                <td>When the number of iterations is unknown</td>
                <td>Can check condition at the start (<code>Do While</code>) or the end (<code>Loop While</code>).</td>
            </tr>
        </tbody>
    </table>


### 1. For..Next
<h6>loop is used when you know the number of iterations in advance.</h6>
<i>Syntax : 

```bash 
For counter = start To end [Step increment]
    ' Code to execute
Next
```
<br>
</i>   
<i>Example 1 : 

```bash 
Dim i
For i = 1 To 5
    WScript.Echo "Iteration: " & i
Next

```
</i>
</i>   
<i>Output : 

```bash 
Iteration: 1
Iteration: 2
Iteration: 3
Iteration: 4
Iteration: 5
```
</i><br>
<i> Example 2 : 

```bash 
For i = 10 To 1 Step -2
    WScript.Echo i
Next

```
</i>
<i>Output : 

```bash 
Iteration: 10
Iteration: 8
Iteration: 6
Iteration: 4
Iteration: 2
```
</i><br>

### 2. For Each...Next
<h6>loop is used to iterate through all items in a collection or array.</h6>

<i> Syntax : 
```bash 
For Each item In collection
    ' Code to execute'
Next
```
</i><br>

<i> Example 1 : 
```bash 
Dim arr, item
arr = Array(10, 20, 30)

For Each item In arr
    WScript.Echo "Value: " & item
Next

```
</i>

<i> Output : 
```bash 
Value: 10
Value: 20
Value: 30
```
</i><br>

### 3. Do...Loop
<h6>The Do...Loop is used when the number of iterations is not known in advance. It can be controlled by a condition at the beginning (Do While or Do Until) or the end (Loop While or Loop Until).</h6>

<i> Syntax : 
```bash 
Do While condition
    ' Code to execute'
Loop

'or'

Do
    ' Code to execute'
Loop While condition

```
</i><br>

<i> Example Do While : 
```bash 
Dim x
x = 0

Do While x < 5
    WScript.Echo "Value of x: " & x
    x = x + 1
Loop
```
</i>

<i> Output : 
```bash 
Value of x: 0
Value of x: 1
Value of x: 2
Value of x: 3
Value of x: 4

```
</i><br>
<i> Example Do Until : 
```bash 
Dim y
y = 0

Do Until y = 5
    WScript.Echo "Value of y: " & y
    y = y + 1
Loop

```
</i>

<i> Output : 
```bash 
Value of y: 0
Value of y: 1
Value of y: 2
Value of y: 3
Value of y: 4

```
</i><br>
 
<h1 id="Function">Functions</h1>

 <h3>Comparison: Function vs Sub</h3>
    <table>
        <thead>
            <tr>
                <th>Aspect</th>
                <th>Function</th>
                <th>Sub</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td>Returns Value</td>
                <td>Yes</td>
                <td>No</td>
            </tr>
            <tr>
                <td>Call Syntax</td>
                <td>Can be used in expressions.</td>
                <td>Cannot be used in expressions.</td>
            </tr>
            <tr>
                <td>Purpose</td>
                <td>Used to perform calculations or return data.</td>
                <td>Used to perform tasks.</td>
            </tr>
        </tbody>
    </table>

<br>

<i> Syntax of Function
```bash
Function FunctionName([argument1, argument2, ...])
    ' Code block'
    FunctionName = [ReturnValue] ' Return value'
End Function

```
Syntax of Sub
```bash
Sub SubName([argument1, argument2, ...])
    ' Code block'
End Sub
```
</i>
  <table>
        <thead>
            <tr>
                <th>Component</th>
                <th>Description</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td><code>Function Name()</code></td>
                <td>Declares a function with the specified name.</td>
            </tr>
            <tr>
                <td><code>Arguments</code></td>
                <td>Parameters passed to the function.</td>
            </tr>
            <tr>
                <td><code>Return Value</code></td>
                <td>Assign the result to the function name.</td>
            </tr>
            <tr>
                <td><code>End Function</code></td>
                <td>Ends the function definition.</td>
            </tr>
        </tbody>
    </table><br>

<i> Example:
<h6>1. Function with Multiple Arguments</h6>
<b>Function </b>

```bash
Function AddNumbers(a, b)
    AddNumbers = a + b ' Return the sum'
End Function

Dim total
total = AddNumbers(3, 7)
WScript.Echo "The sum is " & total
```
<b>Sub </b>

```bash
Sub CalculateArea(Length, Width)
    Dim Area
    Area = Length * Width
    WScript.Echo "The area of the rectangle is " & Area
End Sub

' Call the Sub with two arguments'
CalculateArea 5, 10

```



<br><h6>2. Function Without Arguments</h6>
<b>Function </b>
```bash
Function GetCurrentDate()
    GetCurrentDate = Date
End Function

WScript.Echo "Today's date is " & GetCurrentDate()

```
<b>Sub </b>
```bash
Sub SayHello()
    WScript.Echo "Hello, World!"
End Sub

' Call the Sub'
SayHello()
```

<br><h6>3.Function with Conditional Logic</h6>
<b>Function </b>
```bash
Function IsEven(Number)
    If Number Mod 2 = 0 Then
        IsEven = True
    Else
        IsEven = False
    End If
End Function

Dim number
number = 8
If IsEven(number) Then
    WScript.Echo number & " is even"
Else
    WScript.Echo number & " is odd"
End If

```
<b>Sub </b>
```bash
Sub DisplayMessage(Message)
    If IsEmpty(Message) Then
        Message = "Default Message"
    End If
    WScript.Echo Message
End Sub

' Call the Sub with and without arguments'
DisplayMessage "Custom Message"


```

<br><h6>4. Advanced: Recursion</h6>
<b>Function</b>
```bash
Function Factorial(n)
    If n = 0 Then
        Factorial = 1
    Else
        Factorial = n * Factorial(n - 1)
    End If
End Function

Dim fact
fact = Factorial(5)
WScript.Echo "5! = " & fact
```
</i>    




<br>
<hr>


### Common VBScript Functions 

 <table>
        <thead>
            <tr>
                <th>Function</th>
                <th>Description</th>
                <th>Example</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td><code>Len(string)</code></td>
                <td>Returns the length of a string.</td>
                <td><code>Len("Hello")</code> = <code>5</code></td>
            </tr>
            <tr>
                <td><code>Left(string, length)</code></td>
                <td>Returns the leftmost characters of a string.</td>
                <td><code>Left("Hello", 2)</code> = <code>"He"</code></td>
            </tr>
            <tr>
                <td><code>Right(string, length)</code></td>
                <td>Returns the rightmost characters of a string.</td>
                <td><code>Right("Hello", 3)</code> = <code>"llo"</code></td>
            </tr>
            <tr>
                <td><code>Mid(string, start, length)</code></td>
                <td>Returns a substring from a string starting at a specified position.</td>
                <td><code>Mid("Hello", 2, 3)</code> = <code>"ell"</code></td>
            </tr>
            <tr>
                <td><code>UCase(string)</code></td>
                <td>Converts a string to uppercase.</td>
                <td><code>UCase("Hello")</code> = <code>"HELLO"</code></td>
            </tr>
            <tr>
                <td><code>LCase(string)</code></td>
                <td>Converts a string to lowercase.</td>
                <td><code>LCase("Hello")</code> = <code>"hello"</code></td>
            </tr>
            <tr>
                <td><code>Trim(string)</code></td>
                <td>Removes leading and trailing spaces from a string.</td>
                <td><code>Trim(" Hello ")</code> = <code>"Hello"</code></td>
            </tr>
            <tr>
                <td><code>IsNumeric(expression)</code></td>
                <td>Checks if an expression is a numeric value.</td>
                <td><code>IsNumeric("123")</code> = <code>True</code></td>
            </tr>
            <tr>
                <td><code>Replace(string, find, replace)</code></td>
                <td>Replaces occurrences of a substring within a string.</td>
                <td><code>Replace("Hello World", "World", "VBScript")</code> = <code>"Hello VBScript"</code></td>
            </tr>
            <tr>
                <td><code>Now()</code></td>
                <td>Returns the current date and time.</td>
                <td><code>Now()</code> = Current Date & Time</td>
            </tr>
            <tr>
                <td><code>Rnd()</code></td>
                <td>Returns a random number between 0 and 1.</td>
                <td><code>Rnd()</code> = Random Number</td>
            </tr>
        </tbody>
    </table>
