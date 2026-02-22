## Add simple content to website
  
To add some simple content to the website use the following steps:  

**Step 1:**  
Add a button to the form.  

**Step 2:**  
Add the following code to the Click event:  
````VBA
    Dim strScript As String
    Static i As Integer
    
    i = i + 1
    strScript = "document.body.innerHTML += '<p>This line was added via VBA (" & i & ")</p>';"
    
    ctlEdgeBrowser.ExecuteJavascript strScript
````

**Explaination VBA-Code:**  
First, we add the necessary JavaScript code to the variable *strScript* to add simple content to the website.  
This code is then executed using the ExecuteJavaScript command.  

**Explaination JavaScript-Code:**
````JavaScript
document.body.innerHTML += '<p>This line was added via VBA (' & i & ')</p>'
````
**document.body**  
Refers to the <body> element of the current webpage.

**innerHTML**  
Represents all HTML content inside the <body>.

**+=**  
Means: add this new HTML to the existing content instead of replacing it.

**'&lt;p&gt;This line was added via VBA (' & i & ')&lt;/p&gt;'**  
This is the HTML string being added.
In your VBA code, the variable i is inserted into the string, so each click adds a numbered line.
