using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;

namespace RetrievePrivateAppointmentsAsDelegate
{
    class Program
    {
        static ExchangeService _exchange = null;
        static string _username = "";
        static string _password = "";
        static string _mailbox = "";

        static void Main(string[] args)
        {
            if (args.Length < 3)
            {
                Console.WriteLine($"Syntax: {System.Reflection.Assembly.GetExecutingAssembly().GetName()} <Username> <Password> <Mailbox> <Date (optional)>");
                return;
            }
            _username = args[0];
            _password = args[1];
            _mailbox = args[2];

            DateTime queryDate = DateTime.Now;
            if (args.Length > 2)
            {
                queryDate = DateTime.Parse(args[3]);
            }

            InitExchange();
            Console.WriteLine($"Reading appointments for date: {queryDate}");
            GetDailyAppointments(queryDate.Date);
            Console.WriteLine("Finished");
            Console.ReadLine();
        }

        static void InitExchange()
        {
            _exchange = new ExchangeService(ExchangeVersion.Exchange2016);
            _exchange.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");
            _exchange.Credentials = new WebCredentials(_username, _password);
            System.IO.File.Delete("trace.log");
            _exchange.TraceListener = new TraceListener("trace.log");
            _exchange.TraceFlags = TraceFlags.All;
            _exchange.TraceEnabled = true;
        }

        static PropertySet RequiredProps()
        {
            // Create and return the PropertySet that we want to retrieve

            return new PropertySet(BasePropertySet.IdOnly,
                ItemSchema.Subject,
                ItemSchema.DateTimeReceived,
                AppointmentSchema.Start,
                AppointmentSchema.End,
                AppointmentSchema.IsAllDayEvent,
                AppointmentSchema.Duration);
        }

        static FolderId SharedCalendarFolder()
        {
            // Return FolderId for the shared mailbox calendar
            return new FolderId(WellKnownFolderName.Calendar, new Mailbox(_mailbox));
        }

        static void GetDailyAppointments(DateTime day)
        {
            // Use CalendarView to retrieve the appointments for the given day

            CalendarView calendarView = new CalendarView(day, day.AddHours(24));
            calendarView.PropertySet = RequiredProps();

            try
            {
                // Read all the appointments and summarise them to console
                FindItemsResults<Appointment> results = _exchange.FindAppointments(SharedCalendarFolder(), calendarView);
                foreach (Appointment appointment in results.Items)
                    Console.WriteLine($"{appointment.Start}: {appointment.Subject}");
            }
            catch (Exception ex)
            {
                if (ex.Message.Equals("The specified object was not found in the store., Item not found."))
                {
                    // This error can occur when a modified occurrence of a private appointment is in the time range
                    // In this case, we need to use an alternative method to retrieve the data
                    Console.WriteLine("Testing for private items within time range (alternate check due to error)");
                    GetDailyAppointmentsWithPrivateModifiedOccurrence(day);
                }
            }
        }

        static void GetDailyAppointmentsWithPrivateModifiedOccurrence(DateTime day)
        {
            // Delegate access to recurring private items fails when a private item has a modified occurrence
            // In this case, we need to find the daily appointments the hard way

            int pageSize = 500;
            ItemView itemView = new ItemView(pageSize, 0, OffsetBasePoint.Beginning);
            itemView.PropertySet = RequiredProps();
            itemView.PropertySet.Add(AppointmentSchema.IsRecurring);
            itemView.PropertySet.Add(new ExtendedPropertyDefinition(DefaultExtendedPropertySet.Appointment, 33329, MapiPropertyType.Integer));
            bool moreItemsAvailable = true;

            try
            {
                while (moreItemsAvailable)
                {
                    FindItemsResults<Item> results = _exchange.FindItems(SharedCalendarFolder(), itemView);
                    moreItemsAvailable = results.MoreAvailable;
                    foreach (Appointment appointment in results.Items)
                        if (appointment.Start.Date.Equals(day.Date))
                            Console.WriteLine($"{appointment.Start}: {appointment.Subject}");
                        else
                        {
                            if (appointment.ExtendedProperties.Count > 0)
                            {
                                bool hasRecurrence = false;
                                foreach (ExtendedProperty extendedProp in appointment.ExtendedProperties)
                                    if (extendedProp.PropertyDefinition.Id == 33329 && (int)extendedProp.Value > 0)
                                        hasRecurrence = true;
                                if (hasRecurrence)
                                {
                                    // We only need to check modified occurrences, as if they are not modified then our original FindItem call would have worked
                                    Console.WriteLine("Checking occurrences");
                                    PropertySet props = RequiredProps();
                                    props.Add(AppointmentSchema.ModifiedOccurrences);
                                    Appointment recurringAppointment = Appointment.Bind(_exchange, appointment.Id, props);
                                    if (recurringAppointment.ModifiedOccurrences.Count > 0)
                                        foreach (OccurrenceInfo info in recurringAppointment.ModifiedOccurrences)
                                            if (info.Start >= day && info.Start < day.AddHours(24))
                                                // This modified occurrence occurs during our time range
                                                Console.WriteLine($"{info.Start}: {recurringAppointment.Subject}");
                                }
                            }
                        }
                    if (moreItemsAvailable)
                        itemView.Offset += pageSize;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
