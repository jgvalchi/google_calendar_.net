using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.IO;
using System.Threading;
using System.Diagnostics;
using System.Globalization;

namespace Google_Calendar
{
    public partial class Menu_Calendario : Form
    {

        static string[] Scopes = { CalendarService.Scope.Calendar };
        static string ApplicationName = "GuateFuturo";
        IFormatProvider provider = new CultureInfo("en-US");

        int cLeft = 1;

        public Menu_Calendario()
        {
            InitializeComponent();
            dateTimePicker1.Format = DateTimePickerFormat.Custom;
            dateTimePicker1.CustomFormat = "dd-MM-yyyy HH:mm";
            dateTimePicker2.Format = DateTimePickerFormat.Custom;
            dateTimePicker2.CustomFormat = "dd-MM-yyyy HH:mm";
            button3.Enabled = false;
            button4.Enabled = false;
        }        
        //boton listar
        private void button1_Click(object sender, EventArgs e)
        {
            Listar_Eventos();
        }
        //boton agregar
        private void button2_Click(object sender, EventArgs e)
        {
            string tmp = textBox1.Text;
            string tmp2 = textBox2.Text;
            string titulo = textBox3.Text;
            string desc = textBox4.Text;
            string strdt = dateTimePicker1.Text;
            DateTime dt = Convert.ToDateTime(dateTimePicker1.Text);
            string strdt2 = dateTimePicker2.Text;
            DateTime dt2 = Convert.ToDateTime(dateTimePicker2.Text);
            
            //mostrando
            //MessageBox.Show("Primero DT-->" + strdt);
            //MessageBox.Show("Segundo DT-->" + strdt2);
            //MessageBox.Show("Primer Textbox -->" + tmp);
            //MessageBox.Show("Segundo Textbox -->" + tmp2);
            //MessageBox.Show("Titulo Textbox -->" + titulo);
            //MessageBox.Show("Descripcion Textbox -->" + desc);

            Agregar_Evento(titulo, desc ,strdt, strdt2, tmp, tmp2);  
        }
        //boton eliminar
        private void button3_Click(object sender, EventArgs e)
        {
            //String id_E= sender.ToString();
            string tmp = comboBox1.Text;
            MessageBox.Show("combobox--> " + tmp);
            //MessageBox.Show("boton 3--> " + id_E);
            Eliminar_Evento(tmp);
        }
        //boton modificar
        private void button4_Click(object sender, EventArgs e)
        {
            string tmp = textBox1.Text;
            string tmp2 = textBox2.Text;
            string tmp3 = comboBox1.Text;
            string titulo = textBox3.Text;
            string desc = textBox4.Text;
            string strdt = dateTimePicker1.Text;
            DateTime dt = Convert.ToDateTime(dateTimePicker1.Text);
            string strdt2 = dateTimePicker2.Text;
            DateTime dt2 = Convert.ToDateTime(dateTimePicker2.Text);
            MessageBox.Show("combobox--> " + tmp3);
            Modificar_Evento(titulo, desc, strdt, strdt2, tmp, tmp2, tmp3);            
        }
        //boton agregar email
        private void button5_Click(object sender, EventArgs e)
        {
            AddNewTextBox();
        }

        private void Listar_Eventos()
        { 
           UserCredential credential;
            
            using (var stream =
                new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, "GuateFuturo");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                //Console.WriteLine("Credential file saved to: " + credPath);
                //MessageBox.Show("Credential file saved to: " + credPath);
            }

            // Create Google Calendar API service.
            var service = new CalendarService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            //limpiar combobox
            comboBox1.Items.Clear();
            
            // Define parameters of request.
            EventsResource.ListRequest request = service.Events.List("primary");
            request.TimeMin = DateTime.Now;
            request.ShowDeleted = false;
            request.SingleEvents = true;
            request.MaxResults = 10;
            request.OrderBy = EventsResource.ListRequest.OrderByEnum.StartTime;

            // List events.
            Events events = request.Execute();
            Console.WriteLine("Upcoming events:");
            if (events.Items != null && events.Items.Count > 0)
            {
                foreach (var eventItem in events.Items)
                {
                    string when = eventItem.Start.DateTime.ToString();
                    string id_W = eventItem.Id;
                    if (String.IsNullOrEmpty(when))
                    {
                        when = eventItem.Start.Date;
                    }
                    //Console.WriteLine("{0} ({1})", eventItem.Summary, when);
                    MessageBox.Show(eventItem.Summary +" -- "+ when + " -- "+ id_W);
                    //comboBox1.Items.Add(eventItem.Summary + " -- " + when);
                    comboBox1.Items.Add(id_W);
                    //button3_Click(id_W,null);
                    //MessageBox.Show(when);
                }
                button3.Enabled = true;
                button4.Enabled = true;
            }
            else
            {
                //Console.WriteLine("No upcoming events found.");
                button3.Enabled = false;
                button4.Enabled = false;
                comboBox1.ResetText();
                MessageBox.Show("No upcoming events found.");
            }
                      
            Console.Read();
        }

        private void Agregar_Evento(String Titulo, String Desc, String FI, String FF, String E1, String E2)
        {
            String FIT = FI;
            String FFT = FF;
            String E1T = E1;
            String E2T = E2;
            String TituloT = Titulo;
            String DescT = Desc;
            //String ID_ET = ID_E;
            //MessageBox.Show("Primero DT-->" + FIT);
            //MessageBox.Show("Segundo DT-->" + FFT);
            UserCredential credential;

            using (var stream =
                new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, "GuateFuturo");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                //Console.WriteLine("Credential file saved to: " + credPath);
                //MessageBox.Show("Credential file saved to: " + credPath);
            }

            // Create Google Calendar API service.
            var service = new CalendarService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            //Nueve evento
            Event newEvent = new Event()
            {
                //Summary = TituloT,
                Summary = "Titulo",
                Location = "800 Howard St., San Francisco, CA 94103",
                //Description = DescT,
                Description = "Descripcion",
                //Id = id_generado,
                Start = new EventDateTime()
                {
                    DateTime =  Convert.ToDateTime(FIT),
                    TimeZone = "America/Guatemala",
                },
                End = new EventDateTime()
                {
                    DateTime = Convert.ToDateTime(FFT),
                    TimeZone = "America/Guatemala",
                },
                Recurrence = new String[] { "RRULE:FREQ=DAILY;COUNT=1" },
                Attendees = new EventAttendee[] {
                    //new EventAttendee() { Email = E1T },
                    //new EventAttendee() { Email = E2T },
                     new EventAttendee() { Email = "jgvalchi@gmail.com" },
                    new EventAttendee() { Email = "jgvalchigame@gmail.com" },
                },
                Reminders = new Event.RemindersData()
                {
                    UseDefault = false,
                    Overrides = new EventReminder[] {
                    new EventReminder() { Method = "email", Minutes = 24 * 60 },
                    new EventReminder() { Method = "sms", Minutes = 10 },
                    }
                }
            };

            String calendarId = "primary";
            EventsResource.InsertRequest request = service.Events.Insert(newEvent, calendarId);
            Event createdEvent = request.Execute();
            //Console.WriteLine("Event created: {0}", createdEvent.HtmlLink);
            MessageBox.Show("Evento creado Link --> {0}" + createdEvent.HtmlLink);
            Console.Read();
            Listar_Eventos();
        }

        private void Modificar_Evento(String Titulo, String Desc, String FI, String FF, String E1, String E2, String event_Id)
        {
            String FIT = FI;
            String FFT = FF;
            String E1T = E1;
            String E2T = E2;
            String TituloT = Titulo;
            String DescT = Desc;
            String eventId = event_Id;
            //String ID_ET = ID_E;
            //MessageBox.Show("Primero DT-->" + FIT);
            //MessageBox.Show("Segundo DT-->" + FFT);
            UserCredential credential;

            using (var stream =
                new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, "GuateFuturo");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                //Console.WriteLine("Credential file saved to: " + credPath);
                //MessageBox.Show("Credential file saved to: " + credPath);
            }

            // Create Google Calendar API service.
            var service = new CalendarService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            //Modificar evento
            Event updateEvent = new Event()
            {
                //Summary = TituloT,
                Summary = "Titulo2eventomodificado",
                Location = "800 Howard St., San Francisco, CA 94103",
                //Description = DescT,
                Description = "Descripcion",
                //Id = id_generado,
                Start = new EventDateTime()
                {
                    DateTime = Convert.ToDateTime(FIT),
                    TimeZone = "America/Guatemala",
                },
                End = new EventDateTime()
                {
                    DateTime = Convert.ToDateTime(FFT),
                    TimeZone = "America/Guatemala",
                },
                Attendees = new EventAttendee[] {
                    //new EventAttendee() { Email = E1T },
                    //new EventAttendee() { Email = E2T },
                     new EventAttendee() { Email = "jgvalchi@gmail.com" },
                    new EventAttendee() { Email = "jgvalchigame@gmail.com" },
                },
                Reminders = new Event.RemindersData()
                {
                    UseDefault = false,
                    Overrides = new EventReminder[] {
                    new EventReminder() { Method = "email", Minutes = 24 * 60 },
                    new EventReminder() { Method = "sms", Minutes = 10 },
                    }
                }
            };

            String calendarId = "primary";
            EventsResource.UpdateRequest request2 = service.Events.Update(updateEvent, calendarId, eventId);
            request2.Execute();
            //Console.WriteLine("Event created: {0}", createdEvent.HtmlLink);
            MessageBox.Show("Evento Modificado");
            Console.Read();
        }

        private void Eliminar_Evento(String id_evento)
        {
            String id_eventoT = id_evento;
            UserCredential credential;
            MessageBox.Show("eliminar --> " + id_eventoT);

            using (var stream =
                new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, "GuateFuturo");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                //Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Google Calendar API service.
            var service = new CalendarService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            //Eliminar evento
            EventsResource.DeleteRequest request3 = service.Events.Delete("primary", id_eventoT);
            String deleteEvent2 = request3.Execute();
            MessageBox.Show("Evento Eliminado");
            Console.Read();
            Listar_Eventos();
        }

        public System.Windows.Forms.TextBox AddNewTextBox()
        {
            System.Windows.Forms.TextBox txt = new System.Windows.Forms.TextBox();
            
            this.Controls.Add(txt);
            int margen_top = cLeft * 25;
            txt.Top = margen_top + 151;
            txt.Left = 176;
            //txt.Text = "TextBox " + this.cLeft.ToString();
            cLeft = cLeft + 1;
            return txt;
        }
    }
}





