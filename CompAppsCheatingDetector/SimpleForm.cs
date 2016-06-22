using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CompAppsCheatingDetector {
    class SimpleForm {

        public string[] ShowDialog() {

            Form prompt = new Form() {

                Width = 500,
                Height = 350,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                Text = "Prompt",
                StartPosition = FormStartPosition.CenterScreen
            };

            prompt.TopMost = true;
           
            string[] results = new string[4];

            Label assignLabel = new Label() { Left = 50, Top = 20, Text = "Assignment Name:" };
            TextBox assignBox = new TextBox() { Left = 50, Top = 50, Width = 400 };
            Label courseLabel = new Label() { Left = 50, Top = 80, Text = "Course Code:" };
            TextBox courseBox = new TextBox() { Left = 50, Top = 110, Width = 400 };
            Label profLabel = new Label() { Left = 50, Top = 140, Text = "Professor:" };
            TextBox profBox = new TextBox() { Left = 50, Top = 170, Width = 400 };
            Label taLabel = new Label() { Left = 50, Top = 200, Text = "Teaching Assistant:" };
            TextBox taBox = new TextBox() { Left = 50, Top = 230, Width = 400 };

            Button confirmation = new Button() { Text = "Ok", Left = 350, Width = 100, Top = 260, DialogResult = DialogResult.OK };
            confirmation.Click += (sender, e) => {
                prompt.Close();
                results[0] = assignBox.Text;
                results[1] = courseBox.Text;
                results[2] = profBox.Text;
                results[3] = taBox.Text;
            };
            prompt.Controls.Add(assignBox);
            prompt.Controls.Add(courseBox);
            prompt.Controls.Add(profBox);
            prompt.Controls.Add(taBox);
            prompt.Controls.Add(assignLabel);
            prompt.Controls.Add(courseLabel);
            prompt.Controls.Add(profLabel);
            prompt.Controls.Add(taLabel);
            prompt.Controls.Add(confirmation);
            prompt.AcceptButton = confirmation;

            return prompt.ShowDialog() == DialogResult.OK ? results : null;
        }
    }
}
