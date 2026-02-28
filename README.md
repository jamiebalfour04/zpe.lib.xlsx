<h1>zpe.lib.xlsx</h1>

<p>
  This is the official XLSX plugin for ZPE.
</p>

<p>
  The plugin provides the ability to create, open, modify and save Microsoft Excel (.xlsx) files directly from ZPE.
</p>

<h2>Installation</h2>

<p>
  Place <strong>zpe.lib.xlsx.jar</strong> in your ZPE native-plugins folder and restart ZPE.
</p>

<p>
  You can also download with the ZULE Package Manager by using:
</p>
<p>
  <code>zpe --zule install zpe.lib.xlsx.jar</code>
</p>

<h2>Documentation</h2>

<p>
  Full documentation, examples and API reference are available here:
</p>

<p>
  <a href="https://www.jamiebalfour.scot/projects/zpe/documentation/plugins/zpe.lib.xlsx/" target="_blank">
    View the complete documentation
  </a>
</p>

<h2>Example</h2>

<pre>

import "zpe.lib.xlsx"

wb = xlsx_new()

sheet = wb->get_sheet(0)

sheet->set_cell(0, 0, "Name")
sheet->set_cell(0, 1, "Age")

sheet->set_cell(1, 0, "Jamie")
sheet->set_cell(1, 1, 34)

sheet->set_cell(2, 0, "Alice")
sheet->set_cell(2, 1, 12)

wb->save("example.xlsx")
wb->close()

</pre>

<h2>Notes</h2>

<ul>
  <li>Uses Apache POI internally for Excel file handling.</li>
  <li>Supports creating new workbooks and opening existing .xlsx files.</li>
  <li>Cell values are automatically handled as strings, numbers or booleans.</li>
  <li>File open and save operations require appropriate ZPE permission levels.</li>
  <li>Cross-platform (Windows, macOS, Linux).</li>
  <li>Designed for seamless integration within the ZPE runtime environment.</li>
</ul>
