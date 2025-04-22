📊 Excel Version Comparator (Cyclomatic Complexity Checker)

A simple C# console tool to compare two Excel files (representing old and new versions) based on user-defined key and value columns. It detects updates, flags records requiring a fix based on a numeric threshold, and tracks changes using a custom version label. Optionally, it can split the results into multiple sheets for better analysis.
<hr>
🔧 Features
<br>
<br>
✅ Compare Excel files using custom key and value columns
<br>
✏️ Annotate changes with detailed notes (Before / After)
<br>
🚩 Automatically detect if a fix is required based on a numeric threshold
<br>
🏷 Tag updates with a custom version label
<br>
🗂 Optionally split output Excel into multiple sheets by column value
<br>
💡 Add Fix Required, Fixed?, and Updated in [version] columns for clarity
<hr>
🧑‍💻 How to Use
<br>
<br>
Run the application (via executable or from Visual Studio/.NET CLI)
<br>
<br>
Provide the following inputs when prompted:
<br>
✅ OLD Excel file path
<br>
✅ NEW Excel file path
<br>
✅ Output Excel file path
<br>
✅ Version label (e.g., v1.4.00000)
<br>
✅ Numeric threshold to determine if a fix is needed (e.g., 10)
<br>
✅ Key columns to match records (e.g., A,B)
<br>
✅ Value columns to compare (e.g., G)
<br>
✅ (Optional) Column to split into multiple sheets (e.g., B)
<br><hr>
📂 Output Overview
<br>
<br>
The generated Excel file includes:
<br>
AllData sheet: full dataset with annotated changes
<br>
Updated in [version]: shows whether a row was modified
<br>
Fix Required: Yes if the new value exceeds the threshold
<br>
Fixed?: default is No (can be manually updated later)
<br>
Note: details the value changes (Before / After)
<br>
📄 Optional: Additional sheets if output is split by a specific column
<br>
<hr>
🛠 Requirements
<br>
<br>
.NET 6.0 SDK or later
<br>
EPPlus NuGet package
