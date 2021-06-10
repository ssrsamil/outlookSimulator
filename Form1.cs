using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OpenPop.Pop3;
using OpenPop.Mime;
using Message = OpenPop.Mime.Message;
using System.Configuration;

namespace OutlookSimulator
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
        private void SendEmail_button(object sender, EventArgs e)
        {
            if (textBox7.Text == "")
            {
                MessageBox.Show("Input username/email account");
            }
            if (textBox6.Text == "")
            {
                MessageBox.Show("Input password");
            }
            else
                SendEmailMethod(textBox7.Text, textBox6.Text);
        }

        private void SendEmailMethod(string user, string pass)
        {
            SmtpClient client = new SmtpClient("smtp.gmail.com");
            using (MailMessage mail = new MailMessage())
            {

                mail.From = new MailAddress(user);
                mail.To.Add(textBox2.Text);
                mail.Subject = textBox4.Text;
                mail.Body = textBox3.Text;

                client.Host = "smtp.gmail.com";
                client.Port = 587;
                client.EnableSsl = true;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.UseDefaultCredentials = false;
                client.Credentials = new NetworkCredential(user, pass);
                client.Send(mail);
            }
        }

        DataTable ReadEmailsFromId()
        {
            DataTable table = new DataTable();

            using (Pop3Client client = new Pop3Client())
            {
                client.Connect("pop.gmail.com", 995, true); //For SSL                
                client.Authenticate(textBox7.Text, textBox6.Text, AuthenticationMethod.UsernameAndPassword);

                int messageCount = client.GetMessageCount();
                for (int i = messageCount; i > 0; i--)
                {
                    table.Rows.Add(client.GetMessage(i).Headers.Subject, client.GetMessage(i).Headers.DateSent);
                    string msdId = client.GetMessage(i).Headers.MessageId;
                    OpenPop.Mime.Message msg = client.GetMessage(i);
                    OpenPop.Mime.MessagePart plainTextPart = msg.FindFirstPlainTextVersion();
                    string message = plainTextPart.GetBodyAsText();
                }
            }
            return table;
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            if (Control.IsKeyLocked(Keys.CapsLock))
            {
                MessageBox.Show("The Caps Lock key is ON.");
            }
            if (textBox6.Text == "")
            {
                MessageBox.Show("Input password");
            }

        }

        private void label4_Click(object sender, EventArgs e)
        {
            label4.Text = "enter pass";
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            using (OpenPop.Pop3.Pop3Client client = new OpenPop.Pop3.Pop3Client())
            {
                StringBuilder builder = new StringBuilder();

                client.Connect("pop.gmail.com", 995, true);
                client.Authenticate(textBox7.Text, textBox6.Text);

                textBox5.AppendText("Checking inbox\n");

                var count = client.GetMessageCount();

                if (count == 0)
                {
                    textBox5.AppendText("Mailbox is empty!");
                    return;
                }

                for (int i = 0; i < count; i++)
                {
                    var message = client.GetMessage(i+1);

                    if (!message.Headers.From.MailAddress.Address.Equals(textBox7.Text, StringComparison.InvariantCultureIgnoreCase))
                    {
                        MessagePart plainText = message.FindFirstPlainTextVersion();

                        builder.Append(plainText.GetBodyAsText());
                        textBox5.AppendText(builder.ToString());
                        textBox5.AppendText("\n==============================================================================\n");
                    }
                }
            }
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }

        public string displayMembers(List<Message> m)
        {
            foreach (Message s in m)
            {
                return textBox5.Text = String.Join(Environment.NewLine, m);
            }
            return null;
        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            if (Control.IsKeyLocked(Keys.CapsLock))
            {
                MessageBox.Show("The Caps Lock key is ON.");
            }
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }
    }
}

