Begin Tests
----Add---
Add Cor
Redmond
23
72.5
Add - Missing Key
No Key Data
Add - Duplicate Key
Duplicate Key
Redmond
Test that the ForEach iterators get updated when adding items
1) a
2) a
2) b
2) c
2) d
2) e
2) f
1) b
2) aa
2) a
2) b
2) c
2) d
2) e
2) f
1) c
2) aa
2) aa
2) a
2) b
2) c
2) d
2) e
2) f
1) d
2) aa
2) aa
2) aa
2) a
2) b
2) c
2) d
2) e
2) f
1) e
2) aa
2) aa
2) aa
2) aa
2) a
2) b
2) c
2) d
2) e
2) f
1) f
2) aa
2) aa
2) aa
2) aa
2) aa
2) a
2) b
2) c
2) d
2) e
2) f
----Count---
Count - No Item
0
Count - With Items
1
2
3
----Item---
Item - Existent Index
Redmond
Redmond
23
23
72.5
72.5
Item - Non Existent Index
Missing Index
Missing Key
----Remove---
Remove - Existent Index
Marks.1 Removed
2
City.1 Removed
1
Fraction.1 Removed
0
Remove - Non Existent Index
Error Removing Non Existent Key
Error Removing Non Existent Index
Test that the ForEach iterators get updated when removing items
1) a
2) a
2) b
2) c
2) d
2) e
2) f
1) b
2) b
2) c
2) d
2) e
2) f
1) c
2) c
2) d
2) e
2) f
1) d
2) d
2) e
2) f
1) e
2) e
2) f
1) f
2) f
----IList----
Adding 9 1-based collection items
      Item(1)=Data1
      Item(2)=Data2
      Item(3)=Data3
      Item(4)=Data4
      Item(5)=Data5
      Item(6)=Data6
      Item(7)=Data7
      Item(8)=Data8
      Item(9)=Data9
Converting collection to a 0-based list
   Inserting After 2
      Item(0)=Data1
      Item(1)=Data2
      Item(2)=Inserted
      Item(3)=Data3
      Item(4)=Data4
      Item(5)=Data5
      Item(6)=Data6
      Item(7)=Data7
      Item(8)=Data8
      Item(9)=Data9
   IList.RemoveAt(5): passed
   IList.RemoveAt() all items: passed
   IList.Clear (empty list): passed
   IList.Clear: passed
   IList.Add: passed
   IList.Contains: passed
   IList.IndexOf: passed
   IList.Remove: passed
   IList.IsFixedSize: passed
   IList.IsReadOnly: passed
   IList.Item: passed
----ICollection----
ICollection.Count : passed
ICollection.IsSynchronized : passed
ICollection.SyncRoot : passed
ICollection.CopyTo( )
    ICollection.CopyTo() Null argument : passed
    ICollection.CopyTo( New Object(1,1) { }, 0 ) : passed
    ICollection.CopyTo( Object(), -1 ) : passed
    ICollection.CopyTo( Object(), 5 ) : passed
    ICollection.CopyTo( Object(), 1 ) : passed
    ICollection.CopyTo( String(), 1 ) : passed
*** Bug 254032
   1) passed
   2) passed
   3) passed
   4) passed
End Tests
