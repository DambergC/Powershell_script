<B>Script to Create one or more Maintance Windows for a Collection in MECM</b>
<p>
If you need to create one or more Maintance Windows in MECM for a Collection you can use this script.
<p>You will have the following options
<ol>
  <li>CollID - CollectionID
  <li>Offweek - How many weeks after patch tuesday
  <li>Offdays - How many days after patch tuesday
  <li>StartHour - The hour (0-23) when the Maintance Window should start
  <li>StartMinutes - On the minute (0-59) when the Maintance Window should start
  <li>EndHour - How long in hour for the Maintance Window
  <li>EndMinutes - On the minute for the Maintance Window  
  <li>Patchmount - Which month you want the Maintance Window (1-12)
  <li>PatchYear - Which year
  <li>ClearOldMW - If you want to clean up old Maintance Window for the Collection
  <li>ApplyTo - If you want Any, TaskSequence or Only SoftwareUpdates to be controled by Maintance Window  
</ol>    
<p>From your input the script will calculate Patch Tuesday for the month and set start- and stoptime for the maintance Window.
