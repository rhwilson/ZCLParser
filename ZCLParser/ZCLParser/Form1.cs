using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlServerCe;

namespace ZCLParser
{
    public partial class Main : Form
    {
        public Main()
        {
            InitializeComponent();
        }      

        public void ZCLParse()
        {
            // clear all data first
            ClearAllData();

            // set counter to 0
            int counter = 0;

            // create openfiledialog for selecting zcl file
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "Open Zigbee Cluster Library file to parse";
            ofd.Filter = "Zigbee Cluster Library|*.zcl";

            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                // file information
                string result;
                result = Path.GetFileNameWithoutExtension(ofd.FileName);
                labelName.Text = result;

                // get size of file
                var size = new FileInfo(ofd.FileName).Length;
                labelSize.Text = size.ToString() + " kb";

                // create datagridview Columns                
                dataGridViewParse.Columns.Add("ClusterNum", "Cluster");
                dataGridViewParse.Columns.Add("ManufacturerSpecific", "Manufacturer Specific");
                dataGridViewParse.Columns.Add("Direction", "Direction");
                dataGridViewParse.Columns.Add("DisableDefaultResponse", "Default Response");
                dataGridViewParse.Columns.Add("Reserved", "Reserved");
                dataGridViewParse.Columns.Add("ManufacturerCode", "Manufacturer Code");
                dataGridViewParse.Columns.Add("TransactionSequenceNum", "Transaction Sequence Number");
                dataGridViewParse.Columns.Add("CommandID", "Command Identifier");
                dataGridViewParse.Columns.Add("FramePayload", "Frame Payload");

                // create variables
                String line;
                Stream myStream = null;
                string ConvertedBinary = "";
                string ClusterNum = "";
                string ManufacturerSpecific = "";
                string Direction = "";
                string DisableDefaultResponse = "";
                string Reserved = "Bits 5-7";
                string ManufacturerCode = "";
                string TransactionSequenceNum = "";
                string CommandID = "";
                string FramePayload = "";
                string filename = ofd.FileName;

                // set delimeters for looping
                char[] delimiterChars = { '\n', '\r' };

                string[] packets = filename.Split(delimiterChars);

                try
                {
                    if ((myStream = ofd.OpenFile()) != null)
                    {
                        using (StreamReader sr = new StreamReader(myStream))
                        {
                            while ((line = sr.ReadLine()) != null)
                            {
                                foreach (string item in packets)
                                {
                                    // count each line as a cluster
                                    counter += 1;

                                    string FrameControl = line.Substring(0, 2);
                                    ConvertedBinary = Convert.ToString(Convert.ToInt32(FrameControl, 16), 2).PadLeft(8, '0');
                                    string FrameType = ConvertedBinary.Substring(6, 2);

                                    // frame type subfield
                                    switch (FrameType)
                                    {
                                        case "00":
                                            ClusterNum = "0x00: Entire Profile";
                                            break;
                                        case "01":
                                            ClusterNum = "0x01: Cluster";
                                            break;
                                        case "10":
                                            ClusterNum = "0x10: Reserved";
                                            break;
                                        case "11":
                                            ClusterNum = "0x11: Reserved";
                                            break;
                                        default:
                                            ClusterNum = "Error";
                                            break;
                                    }

                                    // manufacturer specific subfield
                                    string isManufacturerSpecificCode = ConvertedBinary.Substring(5, 1);
                                    if (Convert.ToInt32(isManufacturerSpecificCode) == 0)
                                    {
                                        ManufacturerSpecific = "Not included";
                                    }
                                    else
                                    {
                                        ManufacturerSpecific = "Is present";
                                    }

                                    // direction subfield
                                    string direction = ConvertedBinary.Substring(4, 1);
                                    switch (direction)
                                    {
                                        case "0":
                                            Direction = "From client";
                                            break;
                                        case "1":
                                            Direction = "From server";
                                            break;
                                        default:
                                            Direction = "Error"; ;
                                            break;
                                    }

                                    // disable default response subfield
                                    string DefaultResponse = ConvertedBinary.Substring(3, 1);
                                    switch (DefaultResponse)
                                    {
                                        case "0":
                                            DisableDefaultResponse = "Will be returned";
                                            break;
                                        case "1":
                                            DisableDefaultResponse = "Returned only if error";
                                            break;
                                        default:
                                            DisableDefaultResponse = "Error";
                                            break;
                                    }

                                    // manufacturer code field
                                    if (Convert.ToInt32(isManufacturerSpecificCode) == 1)
                                    {
                                        string manufacturerCode = line.Substring(2, 4);
                                        ConvertedBinary = Convert.ToString(Convert.ToInt32(manufacturerCode, 16), 2).PadLeft(16, '0');
                                        ManufacturerCode = ConvertedBinary;
                                    }
                                    else
                                    {
                                        ManufacturerCode = "";
                                    }

                                    // transaction sequence number
                                    string transactionSequenceNumber = "";
                                    if (Convert.ToInt32(isManufacturerSpecificCode) == 1)
                                    {
                                        transactionSequenceNumber = line.Substring(6, 2);
                                    }
                                    else
                                    {
                                        transactionSequenceNumber = line.Substring(2, 2);
                                    }
                                    TransactionSequenceNum = transactionSequenceNumber;

                                    // command identifier field
                                    string CommandIdentitifier = "";
                                    if (Convert.ToInt32(isManufacturerSpecificCode) == 1)
                                    {
                                        CommandIdentitifier = line.Substring(8, 2);
                                    }
                                    else
                                    {
                                        CommandIdentitifier = line.Substring(4, 2);
                                    }
                                    switch (CommandIdentitifier)
                                    {
                                        case "00":
                                            CommandID = "0b00: Read Attributes";
                                            break;
                                        case "01":
                                            CommandID = "0b00: Read Attributes Responses";
                                            break;
                                        case "02":
                                            CommandID = "0b00: Write Attributes";
                                            break;
                                        case "03":
                                            CommandID = "0b00: Write Attributes Undivided";
                                            break;
                                        case "04":
                                            CommandID = "0b00: Write Attributes Response";
                                            break;
                                        case "05":
                                            CommandID = "0b00: Write Attributes No Response";
                                            break;
                                        case "06":
                                            CommandID = "0b00: Configure Reporting";
                                            break;
                                        case "07":
                                            CommandID = "0b00: Configure Reporting Response";
                                            break;
                                        case "08":
                                            CommandID = "0b00: Read Reporting Configuration";
                                            break;
                                        case "09":
                                            CommandID = "0b00: Read Reporting Configuration Response";
                                            break;
                                        case "0A":
                                            CommandID = "0b00: Report Attributes";
                                            break;
                                        case "0B":
                                            CommandID = "0b00: Deafault Response";
                                            break;
                                        case "0C":
                                            CommandID = "0b00: Discover Attributes";
                                            break;
                                        case "0D":
                                            CommandID = "0b00: Discover Attributes Response";
                                            break;
                                        case "0E":
                                            CommandID = "0b00: Read Attributes Structured";
                                            break;
                                        case "0F":
                                            CommandID = "0b00: Write Attributes Structured";
                                            break;
                                        case "10":
                                            CommandID = "0b00: Write Attributes Structured Response";
                                            break;
                                        default:
                                            CommandID = "0b01: Cluster Specific Command";
                                            break;
                                    }

                                    // frame payload field
                                    string fPayload = "";
                                    if (Convert.ToInt32(isManufacturerSpecificCode) == 1)
                                    {
                                        // check to see if there is any payload
                                        if (line.Length > 10)
                                        {
                                            fPayload = line.Remove(0, 9);
                                        }
                                        else
                                        {
                                            fPayload = "None";
                                        }
                                    }
                                    else
                                    {
                                        if (line.Length > 6)
                                        {
                                            fPayload = line.Remove(0, 5);
                                        }
                                        else
                                        {
                                            fPayload = "None";
                                        }
                                    }
                                    FramePayload = fPayload;                                    
                                }

                                // add all parsed packets to dataGridView
                                dataGridViewParse.Rows.Add(ClusterNum, ManufacturerSpecific, Direction, DisableDefaultResponse, Reserved, ManufacturerCode, TransactionSequenceNum, CommandID, FramePayload);

                                // add number of total packets to label
                                labelTtlPackets.Text = counter.ToString();
                            }
                        }
                    }
                }

                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
                }

                for (int i = 0; i < dataGridViewParse.Rows.Count; i++)
                {
                    dataGridViewParse.Rows[i].HeaderCell.Value = (i + 1).ToString();
                }

                dataGridViewParse.AutoResizeColumns();
            }
        }

        public static bool TableExists(SqlCeConnection connection, string tableName)
        {
            // checking for existing table in database (file name)
            using (var command = new SqlCeCommand())
            {
                command.Connection = connection;
                var sql = string.Format("SELECT COUNT(*) FROM information_schema.tables WHERE table_name = '{0}'", tableName);
                command.CommandText = sql;
                var count = Convert.ToInt32(command.ExecuteScalar());
                return (count > 0);
            }
        }

        public void ConditionallyCreateTables()
        {
            // create variables
            string filename = labelName.Text;
            string messageFirst = "Table doesn't exist. Would you like to create it?";
            string captionFirst = "ZCL frame format: Create Table";
            string messageSecond = "File: " + labelName.Text + " was successfully added to the correlating table. " + labelTtlPackets.Text + " packets were added.";
            string captionSecond = "ZCL frame format: Parsed Data";
            const string sdfPath = @"I:\dev\c#\ZCLParser\ZCLParser\Packets.sdf";
            string dataSource = string.Format("Data Source={0}", sdfPath);

            // check if database exists
            //if (!File.Exists(sdfPath))
            //{
            //    using (var engine = new SqlCeEngine(dataSource))
            //    {
            //        engine.CreateDatabase();
            //    }
            //}
            using (var connection = new SqlCeConnection(dataSource))
            {
                connection.Open();
                using (var command = new SqlCeCommand())
                {
                    command.Connection = connection;

                    // check if table exists named from the file
                    if (!TableExists(connection, filename))
                    {
                        var dialogResultCreate = MessageBox.Show(messageFirst, captionFirst, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (dialogResultCreate == DialogResult.Yes)
                        {
                            // create table
                            command.CommandText = "CREATE TABLE " + filename + "(ClusterNum nvarchar(100), ManufacturerSpecific nvarchar(100), Direction nvarchar(100), DisableDefaultResponse nvarchar(100), Reserved nvarchar(100), ManufacturerCode nvarchar(100), "
                    + "TransactionSequenceNum nvarchar(100), CommandID nvarchar(100), FramePayload nvarchar(1000))";

                            try
                            {
                                int count = command.ExecuteNonQuery();

                                if (count < 0)
                                {
                                    MessageBox.Show("Table: " + filename + " created successfully.", "Create New Table: Success", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                                }

                                try
                                {
                                    SqlCeDataAdapter da = new SqlCeDataAdapter();

                                    for (int i = 0; i < dataGridViewParse.Rows.Count - 2; i++)
                                    {
                                        // insert parsed data
                                        String insertData = "INSERT INTO " + filename + "(ClusterNum,ManufacturerSpecific,Direction,DisableDefaultResponse,ManufacturerCode,TransactionSequenceNum,CommandID,FramePayload) VALUES (@ClusterNum,@ManufacturerSpecific,@Direction,@DisableSpecificResponse,@ManufacturerCode,@TransactionSequenceNum,@CommandID,@FramePayload)";
                                        SqlCeCommand cmd = new SqlCeCommand(insertData, connection);
                                        cmd.Parameters.AddWithValue("@ClusterNum", dataGridViewParse.Rows[i].Cells[0].Value);
                                        cmd.Parameters.AddWithValue("@ManufacturerSpecific", dataGridViewParse.Rows[i].Cells[1].Value);
                                        cmd.Parameters.AddWithValue("@Direction", dataGridViewParse.Rows[i].Cells[2].Value);
                                        cmd.Parameters.AddWithValue("@DisableSpecificResponse", dataGridViewParse.Rows[i].Cells[3].Value);
                                        cmd.Parameters.AddWithValue("@ManufacturerCode", dataGridViewParse.Rows[i].Cells[4].Value);
                                        cmd.Parameters.AddWithValue("@TransactionSequenceNum", dataGridViewParse.Rows[i].Cells[5].Value);
                                        cmd.Parameters.AddWithValue("@CommandID", dataGridViewParse.Rows[i].Cells[6].Value);
                                        cmd.Parameters.AddWithValue("@FramePayload", dataGridViewParse.Rows[i].Cells[7].Value);
                                        da.InsertCommand = cmd;
                                        cmd.ExecuteNonQuery();
                                    }

                                    var dialogResultInsert = MessageBox.Show(messageSecond, captionSecond, MessageBoxButtons.OK, MessageBoxIcon.Information);
                                }

                                catch (SqlCeException ex)
                                {
                                    MessageBox.Show("Failed to insert data into new table because: " + ex.Message);
                                }
                                finally
                                {
                                    connection.Close();
                                }
                            }

                            catch (SqlCeException ex)
                            {
                                MessageBox.Show("Failed to create table because: " + ex.Message);
                            }
                            finally
                            {
                                connection.Close();
                            }
                        }

                        else if (dialogResultCreate == DialogResult.No)
                        {
                            // do nothing 
                        }
                    }
                    
                    // file exists
                    else
                    {
                        MessageBox.Show("This file already exists in the database. Please choose another file.", "Create New Table: Failed", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                }
            }
        }

        public void ClearAllData()
        {
            // clear all data to begin another file parsing
            labelName.Text = "";
            labelSize.Text = "";
            labelTtlPackets.Text = "";

            dataGridViewParse.Columns.Clear();
            dataGridViewParse.Rows.Clear();

            treeView1.Nodes.Clear();
        }

        private void toolStripMenuItem_Click(object sender, EventArgs e)
        {
            ZCLParse();
        }

        private void btnOpenFile_Click(object sender, EventArgs e)
        {
            ZCLParse();
        }

        private void btnInsertToDB_Click(object sender, EventArgs e)
        {
            ConditionallyCreateTables();
        }

        public void dataGridViewParse_SelectionChanged(object sender, EventArgs e)
        {
            // clear previous values in nodes
            treeView1.Nodes.Clear();

            // create treeview nodes  
            TreeNode rootNode = new TreeNode("ZigBee ZCL");
            TreeNode zclHeader = new TreeNode("ZCL Header");
            TreeNode zclPayload = new TreeNode("ZCL Payload");
            TreeNode fControl = new TreeNode("Frame Control");
            TreeNode mCode = new TreeNode("Manufacturer Code");
            TreeNode tSeqNum = new TreeNode("Transaction Sequence Number");
            TreeNode cID = new TreeNode("Command Identifier");
            TreeNode fPayload = new TreeNode("Frame Payload");

            // add new nodes
            treeView1.Nodes.Add(rootNode);
            rootNode.Nodes.Add(zclHeader);
            rootNode.Nodes.Add(zclPayload);
            zclHeader.Nodes.Add(fControl);
            zclHeader.Nodes.Add(mCode);
            zclHeader.Nodes.Add(tSeqNum);
            zclHeader.Nodes.Add(cID);
            zclPayload.Nodes.Add(fPayload);

            if (dataGridViewParse.SelectedRows.Count > 0)
            {
                string cluster = dataGridViewParse.SelectedRows[0].Cells["ClusterNum"].Value.ToString();
                string manufacturerSpecific = dataGridViewParse.SelectedRows[0].Cells["ManufacturerSpecific"].Value.ToString();
                string direction = dataGridViewParse.SelectedRows[0].Cells["Direction"].Value.ToString();
                string disableDefaultResponse = dataGridViewParse.SelectedRows[0].Cells["DisableDefaultResponse"].Value.ToString();
                string reserved = dataGridViewParse.SelectedRows[0].Cells["Reserved"].Value.ToString();
                string manufacturerCode = dataGridViewParse.SelectedRows[0].Cells["ManufacturerCode"].Value.ToString();
                string transactionSequenceNum = dataGridViewParse.SelectedRows[0].Cells["TransactionSequenceNum"].Value.ToString();
                string commandID = dataGridViewParse.SelectedRows[0].Cells["CommandID"].Value.ToString();
                string framePayload = dataGridViewParse.SelectedRows[0].Cells["FramePayload"].Value.ToString();

                // add dataGridView values to treeview
                fControl.Nodes.Add(cluster);
                fControl.Nodes.Add(manufacturerSpecific);
                fControl.Nodes.Add(direction);
                fControl.Nodes.Add(disableDefaultResponse);
                fControl.Nodes.Add(reserved);
                mCode.Nodes.Add(manufacturerCode);
                tSeqNum.Nodes.Add(transactionSequenceNum);
                cID.Nodes.Add(commandID);
                fPayload.Nodes.Add(framePayload);
                treeView1.ExpandAll();
            }

            // set colors of nodes
            foreach (TreeNode n in rootNode.Nodes)
            {
                n.ForeColor = Color.OrangeRed;
                if (n.Text.ToString() == "ZCL Payload")
                {
                    n.ForeColor = Color.DodgerBlue;
                }
            }

            foreach (TreeNode n in zclHeader.Nodes)
            {
                n.ForeColor = Color.OrangeRed;
                foreach (TreeNode z in n.Nodes)
                {
                    z.ForeColor = Color.OrangeRed;
                }
            }

            foreach (TreeNode n in zclPayload.Nodes)
            {
                n.ForeColor = Color.DodgerBlue;
                foreach (TreeNode z in n.Nodes)
                {
                    z.ForeColor = Color.DodgerBlue;
                }
            }
        }

        private void dataGridViewParse_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            int firstDisplayedCellIndex = dataGridViewParse.FirstDisplayedCell.RowIndex;
            int lastDisplayedCellIndex = firstDisplayedCellIndex + dataGridViewParse.DisplayedRowCount(true);

            Graphics Graphics = dataGridViewParse.CreateGraphics();
            int measureFirstDisplayed = (int)(Graphics.MeasureString(firstDisplayedCellIndex.ToString(), dataGridViewParse.Font).Width);
            int measureLastDisplayed = (int)(Graphics.MeasureString(lastDisplayedCellIndex.ToString(), dataGridViewParse.Font).Width);

            int rowHeaderWitdh = System.Math.Max(measureFirstDisplayed, measureLastDisplayed);
            dataGridViewParse.RowHeadersWidth = rowHeaderWitdh + 35;
        }        
        
    }
}

