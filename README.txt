A better array for VBA and VBS.

Provides an array-like Class with methods that you'd expect in other languages:

 - First
 - Last
 - LastIndex
 - Length
 - Merge() or Concat()
 - Push()
 - Pop()
 - Shift()
 - Unshift()

Also included in this project is an experiment to bring jQuery-like element search to Word VBA:

 vbaQuery has the following methods:

 - find(selector)

 Eg: Dim Q As New vbaQuery
     Debug.Print Q.find("ROW").Count