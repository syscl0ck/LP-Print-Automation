while ($true) {
    cp .\2025070940025.xlsx .\20250709-tmp.xlsx
    start-sleep -seconds 5
    remove-Item -Path .\20250709-tmp.xlsx
    start-sleep -seconds 5
}