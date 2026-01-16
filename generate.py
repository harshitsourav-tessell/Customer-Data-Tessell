import pandas as pd

# Load your file
file_path = 'query_result_2026-01-16T09_06_51.328485Z.xlsx' # Make sure this matches your file name
df = pd.read_excel(file_path)

# Function to create badges
def create_badge(val, type='default'):
    if pd.isna(val) or str(val).lower() == 'nan':
        return ""
    s_val = str(val)
    if s_val.lower() == 'true':
        return f"<span class='badge-pill badge-success'>{s_val}</span>"
    if s_val.lower() == 'false':
        return f"<span class='badge-pill badge-crit'>{s_val}</span>"
    if s_val.lower() == 'successful':
        return f"<span class='badge-pill badge-success'>{s_val}</span>"
    return f"<span class='badge-pill badge-{type}'>{s_val}</span>"

# Function to safe string
def safe_str(val):
    if pd.isna(val) or str(val).lower() == 'nan':
        return ""
    return str(val)

html_output = ""

# Iterate through ALL rows
for index, row in df.iterrows():
    # MultiDisk Badge
    multidisk_html = create_badge(row['MultiDisk'])
    # ComputeSharing Badge
    sharing_html = create_badge(row['ComputeSharingEnabled'])
    # Encryption Badge
    enc_html = create_badge(row['EnableEncryption'])
    # SSL Badge
    ssl_html = create_badge(row['EnableSSL'])
    # Status Badge
    status_html = create_badge(row['Status'])

    row_html = f"""
                <tr>
                    <td>{safe_str(row['SoftwareImage'])}</td>
                    <td>{safe_str(row['SoftwareImageVersion'])}</td>
                    <td>{multidisk_html}</td>
                    <td>{safe_str(row['Topology'])}</td>
                    <td>{sharing_html}</td>
                    <td class='text-mono'>{safe_str(row['ParameterProfileId'])}</td>
                    <td>{safe_str(row['ComputeType'])}</td>
                    <td class='text-mono'>{safe_str(row['OptionsProfileId'])}</td>
                    <td>{status_html}</td>
                    <td>{safe_str(row['Region'])}</td>
                    <td>{safe_str(row['Storage'])}</td>
                    <td>{enc_html}</td>
                    <td>{ssl_html}</td>
                    <td>{safe_str(row['ScriptId'])}</td>
                    <td>{safe_str(row['ScriptId.1'])}</td>
                    <td class='font-bold text-right'>{row['Count']}</td>
                </tr>"""
    html_output += row_html

# Wrap it in the table structure
full_html = f"""<div class="card card-wide">
    <h3>Detailed Replication Data</h3>
    <div class="list-container" style="overflow-x: auto;">
        <table id="replication-table" class="alert-table" style="min-width: 2200px; border-collapse: collapse;">
            <thead>
                <tr style="background: #f1f3f5; text-align: left;">
                    <th>SoftwareImage</th>
                    <th>SoftwareImageVersion</th>
                    <th>MultiDisk</th>
                    <th>Topology</th>
                    <th>ComputeSharingEnabled</th>
                    <th>ParameterProfileId</th>
                    <th>ComputeType</th>
                    <th>OptionsProfileId</th>
                    <th>Status</th>
                    <th>Region</th>
                    <th>Storage</th>
                    <th>EnableEncryption</th>
                    <th>EnableSSL</th>
                    <th>ScriptId</th>
                    <th>ScriptId.1</th>
                    <th class="text-right">Count</th>
                </tr>
                </thead>
            <tbody class="scroll-area">
{html_output}
            </tbody>
        </table>
    </div>
</div>"""

# Write to file
with open("full_table_rows.html", "w", encoding="utf-8") as f:
    f.write(full_html)

print("Successfully generated HTML with all rows!")