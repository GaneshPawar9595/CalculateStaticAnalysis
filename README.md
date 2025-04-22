Excel Version Comparator (Check Cyclomatic Complexity)
This C# console tool compares two Excel files (representing old and new versions) based on user-defined key and value columns. It identifies updates, flags records requiring fixes (based on a numeric threshold), and supports version tracking. Optionally, it can also split the output into multiple sheets for easier analysis.

🔧 Features
✅ Compare Excel files using custom key and value columns

✏️ Annotate changes with detailed notes (Before/After)

🚩 Auto-detect if a fix is required based on numeric threshold

🏷 Tag updates with a custom version label

🗂 Optional: Split output Excel into multiple sheets by column value

💡 Easy-to-read summary with Fix Required, Fixed?, and version tracking columns

🧑‍💻 How to Use
Run the application (either from the executable or by building from source)

Provide the following inputs when prompted:

Path to the OLD Excel file

Path to the NEW Excel file

Output path for the result Excel file

Version label (e.g., v1.4.00000)

Numeric threshold for detecting issues (e.g., 10)

Key columns to match records (e.g., columns A, B)

Value columns to compare (e.g., column G)

Optional: Column to split results into separate sheets (e.g., column B)

The tool will generate an Excel file with:

✅ Updated rows marked with the provided version

📝 Notes showing changed values

⚠️ Flags for rows exceeding the threshold

🗂 Additional sheets if split was selected

📂 Output Overview
AllData Sheet: Full dataset with annotated changes

Updated in [Version]: Column indicating whether the row was modified

Fix Required: Set to "Yes" if the numeric value exceeds threshold

Fixed?: Defaulted to "No" (can be manually updated later)

Note: Describes the changes (Before/After)

Optional Sheets: Split sheets based on a chosen column

🛠 Requirements
.NET 6 or later

EPPlus NuGet package
➤ Install using:

bash
Copy code
dotnet add package EPPlus
📦 NuGet Dependencies
EPPlus: For reading/writing Excel files

No other dependencies required
