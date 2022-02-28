using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Otpuska
{
    public partial class CustomMonthCalendar : MonthCalendar
    {
        private List<DateTime> holidays = new List<DateTime>();
        private List<DateTime> holidays_tmp = new List<DateTime>();

        private List<DateTime> vacation = new List<DateTime>();
        private List<DateTime> vacation_tmp = new List<DateTime>();

        private List<DateTime> anotherIdDays = new List<DateTime>();//Дни в которые выходит в отпуск сотрудник одновременно с которым нельзя выходить данному сотруднику
        private List<DateTime> anotherIdDays_tmp = new List<DateTime>();

        private List<DateTime> closedDays = new List<DateTime>();
        private List<DateTime> closedDays_tmp = new List<DateTime>();

        public List<DateTime> Vacation { get => vacation; set => vacation = value; }
        public List<DateTime> Holidays { get => holidays; set => holidays = value; }
        public List<DateTime> AnotherIdDays { get => anotherIdDays; set => anotherIdDays = value; }
        public List<DateTime> ClosedDays { get => closedDays; set => closedDays = value; }

        public CustomMonthCalendar()
        {
            InitializeComponent();
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == 0x000F)
            {
                Graphics graphics = Graphics.FromHwnd(this.Handle);
                PaintEventArgs pe = new PaintEventArgs(graphics, new Rectangle(0, 0, this.Width, this.Height));
                OnPaint(pe);

                HitTestInfo hInfo;

                if (Holidays != null)
                {
                    holidays_tmp = new List<DateTime>(Holidays);
                }
                if (Vacation != null)
                {
                    vacation_tmp = new List<DateTime>(Vacation);
                }
                if (AnotherIdDays != null)
                {
                    anotherIdDays_tmp = new List<DateTime>(AnotherIdDays);
                }

                if(closedDays != null)
                {
                    closedDays_tmp = new List<DateTime>(closedDays);
                }

                for (int i = 0; i < Size.Height + 200; i++)// Костыль
                {
                    for (int j = 0; j < Size.Width + 200; j++)// Костыль
                    {
                        hInfo = HitTest(i, j);
                        if(hInfo.HitArea != HitArea.Date)
                        {
                            continue;
                        }
                        if (closedDays_tmp.Contains(hInfo.Time))
                        {
                            //graphics.FillRectangle(new SolidBrush(Color.Black), i + 4, j, 18, 14);
                            Point[] pt = new Point[4];
                            pt[0] = new Point(i + 4, j);
                            pt[1] = new Point(i + 22, j + 14);
                            pt[2] = new Point(i+4, j+14);
                            pt[3] = new Point(i + 22, j);
                            //graphics.DrawLines(new Pen(Color.Black), pt);
                            graphics.DrawLine(new Pen(Color.Black), pt[0], pt[1]);
                            graphics.DrawLine(new Pen(Color.Black), pt[2], pt[3]);
                            closedDays_tmp.Remove(hInfo.Time);
                        }
                        else if (vacation_tmp.Contains(hInfo.Time) && !ClosedDays.Contains(hInfo.Time))
                        {
                            graphics.DrawRectangle(new Pen(Color.Green, 0.3f), i + 4, j, 18, 14);
                            vacation_tmp.Remove(hInfo.Time);
                        }
                        else if (anotherIdDays_tmp.Contains(hInfo.Time) && !ClosedDays.Contains(hInfo.Time))
                        {
                            graphics.DrawRectangle(new Pen(Color.Blue, 0.5f), i + 4, j, 18, 14);
                            anotherIdDays_tmp.Remove(hInfo.Time);
                        }
                        else if (holidays_tmp.Contains(hInfo.Time) && !ClosedDays.Contains(hInfo.Time))
                        {
                            graphics.DrawRectangle(new Pen(Color.Red, 0.1f), i + 4, j, 18, 14);
                            holidays_tmp.Remove(hInfo.Time);
                        }
                    }
                }
            }
            base.WndProc(ref m);
        }

        public void Repaint()
        {
            Message m = new Message();
            m.Msg = 0x000F;
            WndProc(ref m);
        }
    }
}
