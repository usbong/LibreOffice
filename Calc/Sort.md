# LibreOffice Calc: BASIC Computer Language
# Auto-sort based on Target Column

> <b>Sub SortMacro()</b><br/>
> ' https://ask.libreoffice.org/t/looking-for-calc-macro-to-sort-like-excel-can-with-this-one/57270;<br/>
> ' last accessed: 20220723<br/>
> ' answer by: Lupp, 2020-10<br/>
> ' answer by: newbie-02, 2020-10; Excel VBA<br/>
> <br/>
> ' Excel VBA Computer Instructions<br/>
> ' Range("A1:j1000").Sort Key1:=Range("A1")<br/>
> <br/>
> <b>Dim sortDesc(1) As New com.sun.star.beans.PropertyValue </b><br/>
> <b>Dim mySortFields(0) As New com.sun.star.util.SortField </b><br/>
> <br/>
> <b>sheet = ThisComponent.Sheets(0)</b><br/>
> <b>rgN = "A1:j1000"</b> REM range of cells: A1 to J1000<br/>
> <b>targetRg = sheet.getCellRangeByName(rgN)</b><br/>
> <br/>
> <b>sortDesc(0).Name = "IsSortColumns" </b>REM Dispensable since 'False' is default.<br/>
> <b>sortDesc(0).Value = False</b><br/>
> <b>mySortFields(0).Field = 0</b> REM sort based on the first column, i.e. column A<br/>
> <b>mySortFields(0).SortAscending = True</b><br/>
> <b>sortDesc(1).Name = "SortFields"</b><br/>
> <b>sortDesc(1).Value = mySortFields</b><br/>
> <b>targetRg.sort(sortDesc)</b><br/>
> <br/>
> <b>End Sub</b>
