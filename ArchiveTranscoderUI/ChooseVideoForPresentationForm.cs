using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ArchiveTranscoder {
    public partial class ChooseVideoForPresentationForm : Form {
        private StreamGroup videoStreamGroup;

        public StreamGroup ChosenVideoStreamGroup {
            get { return this.videoStreamGroup; }
        }

        private ChooseVideoForPresentationForm() {
            InitializeComponent();
        }

        public ChooseVideoForPresentationForm(TreeView tv) : this() {
            //Copy the tree nodes: just clone the top level.
            foreach (TreeNode cn in tv.Nodes) {
                this.treeView1.Nodes.Add((TreeNode)cn.Clone());
            }
        }

        private void buttonOk_Click(object sender, EventArgs e) {
            if (this.treeView1.SelectedNode != null) {
                TreeNode tn = this.treeView1.SelectedNode;
                if ((tn.Tag != null) && (tn.Tag is StreamGroup)) {
                    StreamGroup sg = (StreamGroup)tn.Tag;
                    if (sg.Payload == "dynamicVideo") {
                        this.videoStreamGroup = sg;
                        this.DialogResult = DialogResult.OK;
                        this.Close();
                        return;
                    }
                }
            }
            MessageBox.Show("Please select a video stream and try again.", "Error: A Video Stream Type Was Not Selected", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void buttonCancel_Click(object sender, EventArgs e) {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void treeView1_DoubleClick(object sender, EventArgs e) {
            if (sender is TreeView) {
                TreeView tv = (TreeView)sender;
                TreeNode tn = tv.SelectedNode;
                if ((tn != null) && (tn.Tag is StreamGroup)) 
				{
                    Participant p = (Participant)tn.Parent.Tag;
                    Conference c = (Conference)tn.Parent.Parent.Tag;
                    StreamGroup s = (StreamGroup)tn.Tag;
                    if (s.Payload == "dynamicVideo") {
                        this.videoStreamGroup = s;
                        this.DialogResult = DialogResult.OK;
                        this.Close();
                        return;
                    }
                }
            }
        }
    }
}
