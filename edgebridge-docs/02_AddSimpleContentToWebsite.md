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

Explaination:  
adsasd
