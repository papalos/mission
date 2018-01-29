using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Forms;

namespace Mission
{
    public partial class NumberNote : Form
    {
        public NumberNote()
        {
            InitializeComponent();
            this.comboBoxBonus1.SelectedIndex = 0;
            this.comboBoxBonus2.SelectedIndex = 0;
        }

        private void btnOnDepertment2_Click(object sender, EventArgs e)
        {
            cBoxDepartment2.Enabled = true;
            cBoxPost2.Enabled = true;
            btnOnDepertment3.Enabled = true;

        }

        private void btnOnDepertment3_Click(object sender, EventArgs e)
        {
            cBoxDepartment3.Enabled = true;
            cBoxPost3.Enabled = true;
            btnOnDepertment4.Enabled = true;
        }

        private void btnOnDepertment4_Click(object sender, EventArgs e)
        {
            cBoxDepartment4.Enabled = true;
            cBoxPost4.Enabled = true;
        }

        private void btnOnCountry2_Click(object sender, EventArgs e)
        {
            cBoxCountry2.Enabled = true;
            tBoxCity2.Enabled = true;
            dateTimePicStart2.Enabled = true;
            dateTimePicEnd2.Enabled = true;
            btnOnCountry3.Enabled = true;
        }

        private void btnOnCountry3_Click(object sender, EventArgs e)
        {
            cBoxCountry3.Enabled = true;
            tBoxCity3.Enabled = true;
            dateTimePicStart3.Enabled = true;
            dateTimePicEnd3.Enabled = true;
        }

        private void btnNext1_Click(object sender, EventArgs e)
        {
            tabControl1.SelectTab(tabPage2);
        }

        private void comboBoxReasonChanged(object sender, EventArgs e)
        {
            if(comboBoxReason.SelectedItem.ToString() == "Иное")
            {
                textBoxReason.Enabled = true;
            }
            else
            {
                textBoxReason.Enabled = false;
            }
        }

        private void btnNext2_Click(object sender, EventArgs e)
        {
            tabControl1.SelectTab(tabPage3);
        }

        private void btnRefer_Click(object sender, EventArgs e)
        {
            /*Шаблонные значения для записки*/
            string[] FindText =
            {
             "%number%", "%data%", "%FIO%", "%department&post_all%", "%country1%", "%city1%", "%start1%",
             "%end1%", "%duration1%", "%country2%", "%city2%", "%start2%", "%end2%", "%duration2%",
             "%country3%", "%city3%", "%start3%", "%end3%", "%duration3%", "%durationTotal%","%org%",
             "%reason%", "%numDocReason%", "%dateDocReason%", "%costMission%","%finSource%","%supperFin%",
             "%bonus1&bonus2%", "%purpose%", "%mission%", "%roadMap%", "%competitionCash%", "%extra%","%itogo%"
            };

            /*******************************************************************************/
            /*Номер справки*/
            string number = textBoxNumberNote.Text;
            /*Дата справки*/
            string data = dateTimePicNow.Text;
            /*ФИО*/
            string fio = FIO.Text;
            /*Места работы и должности*/
            string department_postAll = cBoxDepartmentMain.Text + " - " + cBoxPostMain.Text;
            /*Дополняем поле при необходимости*/
            if (cBoxDepartment2.Enabled == true)
            {
                department_postAll += "^l" + cBoxDepartment2.Text + " - " + cBoxPost2.Text;
            }
            if (cBoxDepartment3.Enabled == true)
            {
                department_postAll += Environment.NewLine + cBoxDepartment3.Text + " - " + cBoxPost3.Text;
            }
            if (cBoxDepartment4.Enabled == true)
            {
                department_postAll += Environment.NewLine + cBoxDepartment4.Text + " - " + cBoxPost4.Text;
            }
            
            /*Время проведенное в первой стране*/
            string country1 = cBoxCountry1.Text;
            string city1 = tBoxCity1.Text;
            string start1 = dateTimePicStart1.Text;
            string end1 = dateTimePicEnd1.Text;
            int intDuration1 = (dateTimePicEnd1.Value.Date - dateTimePicStart1.Value.Date).Days;
            if(cBoxCountry2.Enabled == false) { intDuration1 += 1; }
            string duration1 = intDuration1.ToString();
            System.DateTime dataAllEnd = dateTimePicEnd1.Value.Date;

            /*Время проведенное во второй стране*/
            string country2 = cBoxCountry2.Text;
            string city2 = tBoxCity2.Text;
            string start2;
            string end2;
            string duration2;
            if (cBoxCountry2.Enabled)
            {
                start2 = dateTimePicStart2.Text;
                end2 = dateTimePicEnd2.Text;
                int intDuration2 = (dateTimePicEnd2.Value.Date - dateTimePicStart2.Value.Date).Days;
                if (cBoxCountry3.Enabled == false) { intDuration2 += 1; }
                duration2 = intDuration2.ToString();
                dataAllEnd = dateTimePicEnd2.Value.Date;
            }
            else
            {
                start2 = ""; end2 = ""; duration2 = "";
            }

            Debug.WriteLine(duration2);

            /*Время проведенное в третьей стране*/
            string country3 = cBoxCountry3.Text;
            string city3 = tBoxCity3.Text;
            string start3;
            string end3;
            string duration3;
            if (cBoxCountry3.Enabled)
            {
                start3 = dateTimePicStart3.Text;
                end3 = dateTimePicEnd3.Text;
                duration3 = ((dateTimePicEnd3.Value.Date - dateTimePicStart3.Value.Date).Days + 1).ToString();
                dataAllEnd = dateTimePicEnd3.Value.Date;
            }
            else
            {
                start3 = ""; end3 = ""; duration3 = "";
            }

            /*Общая продолжительность коммандировки*/
            string durationTotal = ((dataAllEnd - dateTimePicStart1.Value.Date).Days + 1).ToString();
            Debug.WriteLine(durationTotal);


            /*Организация*/
            string org = tBoxOrg.Text;
            /*Основания для командировки*/
            string reason = (comboBoxReason.Text!="Иное") ? comboBoxReason.Text : textBoxReason.Text;
            /*Номер основания*/
            string numDocReason = tBoxNumDocReason.Text;
            /*Дата основания*/
            string dateDocReason = dateTimePickDocReason.Text;
            /*Предельная сумма финансирования*/
            string costMission = tBoxCostMission.Text;
            /*Источник финансирования*/
            string finSource = tBoxFinSource.Text;
            /*Превышение финансирования*/
            string supperFin = tBoxSupperFin.Text;
            /*Оплата принимающей стороны*/
            string bonus1_bonus2 = comboBoxBonus1.Text +", " + comboBoxBonus2.Text;
            /*Цель командировки*/
            string purpose = textBoxPurpose.Text;
            /*Обязанности*/
            string mission = textBoxMission.Text;
            /*Дорожная Карта*/
            string roadMap = comboBoxRoadMap.Text;
            /*Направление расходования средств*/
            string competitionCash = comboBoxCompetitionCash.Text;
            Debug.WriteLine(comboBoxCompetitionCash.Text);
            /*Дополнительная информация*/
            string extra = textBoxExtra.Text;

            
            /********************************************************************************************/

            

            /*Шаблоны значений для справки*/
            string[] FindTextRef =
            {
                "%val_1%", "%price_1%", "%duration_1%", "%sum_1%",
                "%val_2%", "%price_2%", "%duration_2%", "%sum_2%",
                "%val_3%", "%price_3%", "%duration_3%", "%sum_3%",
                "%priceDay%", "%durationDay%", "%sumDay%",
                "%valTr%", "%sumTr%",
                "%VairHair%", "%SairHair%",
                "%valOrg%", "%sumOrg%",
                "%valViza%", "%sumViza%",
                "%valMed%", "%sumMed%",
                "%resid1%", "%res1Cost%", "%res1Day%", "%res1Sum%",
                "%resid2%", "%res2Cost%", "%res2Day%", "%res2Sum%",
                "%resid3%", "%res3Cost%", "%res3Day%", "%res3Sum%",
                "%itogo%"
            };

            /***** суточные Первая страна *****/
            //возвращает валюту по названию страны
            string val_1 = (cBoxCountry1.Text == "") ? "" : Methods.getValutaSut(cBoxCountry1.Text);
            double price_1_Int;
            /*Стоимость проживания в стране*/
            //если страна не выбрана стоимость = 0
            if (val_1 == "")
            {
                price_1_Int = 0;
            }
            else
            {
                //иначе если принимающая сторона оплачивает суточные или берет все расходы на себя стоимость так же 0
                if (comboBoxBonus1.Text == "Суточные" || 
                    comboBoxBonus2.Text == "Все расходы" || 
                    comboBoxBonus2.Text == "Проживание, суточные" || 
                    comboBoxBonus2.Text == "Суточные" || 
                    comboBoxBonus2.Text == "Суточные, виза" || 
                    comboBoxBonus2.Text == "Суточные, мед страховка" || 
                    comboBoxBonus2.Text == "Суточные, трансфер")
                {
                    price_1_Int = 0;
                }
                else
                {                    
                    if (cBoxCountry1.Text == "Российская Федерация")
                    {
                        //иначе если едем в РФ и расходы принимающая сторона на себя не берет стоимость одного дня составит 700, если берет то 0
                        price_1_Int = (comboBoxBonus1.Text == "Нет")?700d:0;
                    }
                    else
                    {
                        //если мы не в РФ принимающая сторона ничего не оплачивает берем полуную стоимость проживания
                        //если оплачивает питание, визу, проживание, страховку то лишь 1/3 стоимости
                        //иначе все расходы на принимающей стороне коэф равен 0
                        price_1_Int = (comboBoxBonus1.Text == "Нет")
                                        ? 1d
                                        : (comboBoxBonus1.Text == "Питание" ||
                                            comboBoxBonus1.Text == "Питание, виза" ||
                                            comboBoxBonus1.Text == "Питание, мед страховка" ||
                                            comboBoxBonus1.Text == "Питание, проезд" ||
                                            comboBoxBonus1.Text == "Питание, проживание" ||
                                            comboBoxBonus1.Text == "Питание, трансфер")
                                            ? 0.3
                                            : 0d;
                        //умножаем полученный коэффициент на стоимость одного дня получаем точную стоимость одного дня
                        price_1_Int *= Methods.getSutCash(cBoxCountry1.Text);
                    }
                }
            }
            string price_1_Str = price_1_Int.ToString();
            //копируем продолжительности пребывания в стране из поля справки в поле сметы: суточные
            string duration_1 = duration1;
            //Получаем суммарную стоимость суточных в данной стране
            double sum_1d = price_1_Int * double.Parse(duration_1);
            //Обновляем массив хранящий итоговый результат по каждой валюте
            if (Methods.SumValuta.ContainsKey(val_1))
            {
                Methods.SumValuta[val_1] += sum_1d;
            }
            else
            {
                Methods.SumValuta.Add(val_1,sum_1d);
            }
            //Переводим стоимость суток в строковый формат
            string sum_1 = sum_1d.ToString();

            /****** суточные Вторая страна ********/
            //возвращает валюту по названию страны
            string val_2 = (cBoxCountry2.Text == "") ? "" : Methods.getValutaSut(cBoxCountry2.Text);
            double price_2_Int;
            /*Стоимость проживания в стране*/
            //если страна не выбрана стоимость = 0
            if (val_2 == "")
            {
                price_2_Int = 0;
            }
            else
            {
                //иначе если принимающая сторона оплачивает суточные или берет все расходы на себя стоимость так же 0
                if (comboBoxBonus2.Text == "Все расходы" || 
                    comboBoxBonus1.Text == "Питание" || 
                    comboBoxBonus2.Text == "Проживание, трансфер" || 
                    comboBoxBonus2.Text == "Суточные, виза" ||
                    comboBoxBonus2.Text == "Суточные, мед страховка" ||
                    comboBoxBonus2.Text == "Суточные, трансфер")
                {
                    price_2_Int = 0;
                }
                else
                {
                    if (cBoxCountry2.Text == "Российская Федерация")
                    {
                        //иначе если едем в РФ и расходы принимающая сторона на себя не берет стоимость одного дня составит 700, если берет то 0
                        price_2_Int = (comboBoxBonus1.Text == "Нет") ? 700d : 0;
                    }
                    else
                    {
                        //если мы не в РФ принимающая сторона ничего не оплачивает берем полуную стоимость проживания
                        //если оплачивает питание, визу, проживание, страховку то лишь 1/3 стоимости
                        //иначе все расходы на принимающей стороне коэф равен 0
                        price_2_Int = (comboBoxBonus1.Text == "Нет")
                                        ? 1d
                                        : (comboBoxBonus1.Text == "Питание" ||
                                            comboBoxBonus1.Text == "Питание, виза" ||
                                            comboBoxBonus1.Text == "Питание, мед страховка" ||
                                            comboBoxBonus1.Text == "Питание, проезд" ||
                                            comboBoxBonus1.Text == "Питание, проживание" ||
                                            comboBoxBonus1.Text == "Питание, трансфер")
                                            ? 0.3
                                            : 0d;
                        //умножаем полученный коэффициент на стоимость одного дня получаем точную стоимость одного дня
                        price_2_Int *= Methods.getSutCash(cBoxCountry2.Text);
                    }
                }
            }
            string price_2_Str = price_2_Int.ToString();
            //копируем продолжительности пребывания в стране из поля справки в поле сметы            
            double duration_2d = duration2 == "" ? 0 : double.Parse(duration2);
            string duration_2 = duration2;
            //Получаем суммарную стоимость суточных в данной стране
            double sum_2d = price_2_Int * duration_2d;
            //Обновляем массив хранящий итоговый результат по каждой валюте
            if (Methods.SumValuta.ContainsKey(val_2))
            {
                Methods.SumValuta[val_2] += sum_2d;
            }
            else
            {
                Methods.SumValuta.Add(val_2, sum_2d);
            }
            //Переводим стоимость суток в строковый формат
            string sum_2 = sum_2d.ToString();

            /******* суточные Третья страна *******/
            //возвращает валюту по названию страны
            string val_3 = (cBoxCountry3.Text == "") ? "" : Methods.getValutaSut(cBoxCountry2.Text);
            double price_3_Int;
            /*Стоимость проживания в стране*/
            //если страна не выбрана стоимость = 0
            if (val_3 == "")
            {
                price_3_Int = 0;
            }
            else
            {
                //иначе если принимающая сторона оплачивает суточные или берет все расходы на себя стоимость так же 0
                if (comboBoxBonus2.Text == "Все расходы" ||
                    comboBoxBonus1.Text == "Питание, виза" ||
                    comboBoxBonus2.Text == "Трансфер" ||
                    comboBoxBonus2.Text == "Суточные, мед страховка" ||
                    comboBoxBonus2.Text == "Суточные, трансфер")
                {
                    price_3_Int = 0;
                }
                else
                {
                    if (cBoxCountry3.Text == "Российская Федерация")
                    {
                        //иначе если едем в РФ и расходы принимающая сторона на себя не берет стоимость одного дня составит 700, если берет то 0
                        price_3_Int = (comboBoxBonus1.Text == "Нет") ? 700d : 0;
                    }
                    else
                    {
                        //если мы не в РФ принимающая сторона ничего не оплачивает берем полуную стоимость проживания
                        //если оплачивает питание, визу, проживание, страховку то лишь 1/3 стоимости
                        //иначе все расходы на принимающей стороне коэф равен 0
                        price_3_Int = (comboBoxBonus1.Text == "Нет")
                                        ? 1d
                                        : (comboBoxBonus1.Text == "Питание" ||
                                            comboBoxBonus1.Text == "Питание, виза" ||
                                            comboBoxBonus1.Text == "Питание, мед страховка" ||
                                            comboBoxBonus1.Text == "Питание, проезд" ||
                                            comboBoxBonus1.Text == "Питание, проживание" ||
                                            comboBoxBonus1.Text == "Питание, трансфер")
                                            ? 0.3
                                            : 0d;
                        //умножаем полученный коэффициент на стоимость одного дня получаем точную стоимость одного дня
                        price_2_Int *= Methods.getSutCash(cBoxCountry3.Text);
                    }
                }
            }
            string price_3_Str = price_3_Int.ToString();
            //копируем продолжительности пребывания в стране из поля справки в поле сметы
            double duration_3d = duration3=="" ? 0 : double.Parse(duration3);
            string duration_3 = duration3;
            double sum_3d = price_3_Int * duration_3d;
            //Обновляем массив хранящий итоговый результат по каждой валюте
            if (Methods.SumValuta.ContainsKey(val_3))
            {
                Methods.SumValuta[val_3] += sum_3d;
            }
            else
            {
                Methods.SumValuta.Add(val_3, sum_3d);
            }
            //Переводим стоимость суток в строковый формат
            string sum_3 = sum_3d.ToString();


            /*Рубли день возвращения*/
            //стоимость
            int priceDayInt = int.Parse(durationTotal) == 1 ? 0 : 700;
            string priceDayStr = priceDayInt.ToString(); //в массив
            //продолжительность
            int durationDayInt = int.Parse(durationTotal) == 1 ? 0 : 1;
            string durationDayStr = durationDayInt.ToString();//в массив
            //итого
            int sumDayInt = priceDayInt * durationDayInt;
            
            if (Methods.SumValuta.ContainsKey("Рубли"))
            {
                Methods.SumValuta["Рубли"] += sumDayInt;
            }
            else
            {
                Methods.SumValuta.Add("Рубли", sumDayInt);
            }

            string sumDayStr = sumDayInt.ToString(); //в массив

            /*Дополнительные расходы*/
            //Проезд
            string valTr = comboBoxTransitValuta.Text==""?"": comboBoxTransitValuta.Text;//в массив
            double d_sumTr = textBoxTransitCash.Text == "" ? 0 : double.Parse(textBoxTransitCash.Text);
            if (Methods.SumValuta.ContainsKey(valTr))
            {
                Methods.SumValuta[valTr] += d_sumTr;
            }
            else
            {
                Methods.SumValuta.Add(valTr, d_sumTr);
            }
            string s_sumTr = valTr == "" ? "" : d_sumTr.ToString();//в массив

            //Трансфер аэропорт-отель-аэропрот
            string VairHair = comboBoxTransferValuta.Text == "" ? "" : comboBoxTransferValuta.Text;//в массив
            double d_SairHair = textBoxTransferCash.Text == "" ? 0 : double.Parse(textBoxTransferCash.Text);
            if (Methods.SumValuta.ContainsKey(VairHair))
            {
                Methods.SumValuta[VairHair] += d_SairHair;
            }
            else
            {
                Methods.SumValuta.Add(VairHair, d_SairHair);
            }
            string s_SairHair = VairHair == "" ? "" : d_SairHair.ToString();//в массив

            //Оргвзнос
            string valOrg = comboBoxOrgvznosValuta.Text == "" ? "" : comboBoxOrgvznosValuta.Text;//в массив
            double d_sumOrg = textBoxOrgvznosCash.Text == "" ? 0 : double.Parse(textBoxOrgvznosCash.Text);
            if (Methods.SumValuta.ContainsKey(valOrg))
            {
                Methods.SumValuta[valOrg] += d_sumOrg;
            }
            else
            {
                Methods.SumValuta.Add(valOrg, d_sumOrg);
            }
            string s_sumOrg = valOrg == "" ? "" : d_sumOrg.ToString();//в массив

            //Виза
            string valViza = comboBoxVizaValuta.Text == "" ? "" : comboBoxVizaValuta.Text;//в массив
            double d_sumViza = textBoxVizaCash.Text == "" ? 0 : double.Parse(textBoxVizaCash.Text);
            if (Methods.SumValuta.ContainsKey(valViza))
            {
                Methods.SumValuta[valViza] += d_sumViza;
            }
            else
            {
                Methods.SumValuta.Add(valViza, d_sumViza);
            }
            string s_sumViza = valViza == "" ? "" : d_sumViza.ToString();//в массив

            //Медстраховка
            string valMed = comboBoxMedStrhValuta.Text == "" ? "" : comboBoxMedStrhValuta.Text;//в массив
            double d_sumMed = textBoxMedStrahCash.Text == "" ? 0 : double.Parse(textBoxMedStrahCash.Text);
            if (Methods.SumValuta.ContainsKey(valMed))
            {
                Methods.SumValuta[valMed] += d_sumMed;
            }
            else
            {
                Methods.SumValuta.Add(valMed, d_sumMed);
            }
            string s_sumMed = valMed == "" ? "" : d_sumMed.ToString();//в массив

            /*Проживание страна 1*/
            //возвращает валюту по названию страны
            string resid1 = (cBoxCountry1.Text == "") ? "" : Methods.getValuta(cBoxCountry1.Text);
            double res1CostInt;
            if (resid1 == "")
            {
                res1CostInt = 0;
            }
            else
            {
                //иначе если принимающая сторона оплачивает проживание или берет все расходы на себя 
                //стоимость так же 0
                if (comboBoxBonus1.Text == "Питание, проживание" || 
                    comboBoxBonus2.Text == "Проживание" || 
                    comboBoxBonus2.Text == "Проживание, виза" || 
                    comboBoxBonus2.Text == "Проживание, мед страховка" || 
                    comboBoxBonus2.Text == "Проживание, проезд" || 
                    comboBoxBonus2.Text == "Проживание, суточные" || 
                    comboBoxBonus2.Text == "Проживание, трансфер")
                {
                    res1CostInt = 0;
                }
                else
                {
                    res1CostInt = Methods.getDayCash(cBoxCountry1.Text);
                }
            }
            string res1Cost = res1CostInt.ToString();
            //копируем продолжительности пребывания в стране из поля справки в поле сметы: проживание
            string res1Day = duration1;
            string res1Sum = (resid1 == "") ? "" : (res1CostInt * double.Parse(res1Day)).ToString();
            if (Methods.SumValuta.ContainsKey(resid1))
            {
                if (resid1 != "")
                {
                    Methods.SumValuta[resid1] += double.Parse(res1Sum);
                }
            }
            else
            {
                Methods.SumValuta.Add(resid1, double.Parse(res1Sum));
            }

            /*Проживание страна 2*/
            //возвращает валюту по названию страны
            string resid2 = (cBoxCountry2.Text == "") ? "" : Methods.getValuta(cBoxCountry2.Text);
            double res2CostInt;
            if (resid2 == "")
            {
                res2CostInt = 0;
            }
            else
            {
                //иначе если принимающая сторона оплачивает проживание или берет все расходы на себя 
                //стоимость так же 0
                if (comboBoxBonus1.Text == "Питание, суточные" ||
                    comboBoxBonus2.Text == "Проживание, виза" ||
                    comboBoxBonus2.Text == "Проживание, мед страховка" ||
                    comboBoxBonus2.Text == "Проживание, проезд" ||
                    comboBoxBonus2.Text == "Проживание, суточные" ||
                    comboBoxBonus2.Text == "Проживание, трансфер" ||
                    comboBoxBonus2.Text == "Трансфер")
                {
                    res2CostInt = 0;
                }
                else
                {
                    res2CostInt = Methods.getDayCash(cBoxCountry2.Text);
                }
            }
            string res2Cost = res2CostInt.ToString();
            //копируем продолжительности пребывания в стране из поля справки в поле сметы: проживание
            string res2Day = duration2;
            string res2Sum = (resid2 == "") ? "" : (res2CostInt * double.Parse(res2Day)).ToString();
            if (res2Sum != "")
            {
                if (Methods.SumValuta.ContainsKey(resid2))
                {
                    Methods.SumValuta[resid2] += double.Parse(res2Sum);
                }
                else
                {
                    Methods.SumValuta.Add(resid2, double.Parse(res2Sum));
                }
            }

            /*Проживание страна 3*/
            //возвращает валюту по названию страны
            string resid3 = (cBoxCountry3.Text == "") ? "" : Methods.getValuta(cBoxCountry3.Text);
            double res3CostInt;
            if (resid3 == "")
            {
                res3CostInt = 0;
            }
            else
            {
                //иначе если принимающая сторона оплачивает проживание или берет все расходы на себя 
                //стоимость так же 0
                if (comboBoxBonus1.Text == "Питание, трансфер" ||
                    comboBoxBonus2.Text == "Проживание, мед страховка" ||
                    comboBoxBonus2.Text == "Проживание, проезд" ||
                    comboBoxBonus2.Text == "Проживание, суточные" ||
                    comboBoxBonus2.Text == "Проживание, трансфер" ||
                    comboBoxBonus2.Text == "Трансфер" ||
                    comboBoxBonus2.Text == "Суточные")
                {
                    res3CostInt = 0;
                }
                else
                {
                    res3CostInt = Methods.getDayCash(cBoxCountry3.Text);
                }
            }
            string res3Cost = res3CostInt.ToString();
            //копируем продолжительности пребывания в стране из поля справки в поле сметы: проживание
            string res3Day = duration3;
            string res3Sum = (resid3 == "") ? "" : (res3CostInt * double.Parse(res3Day)).ToString();
            if(res3Sum!="")
            if (Methods.SumValuta.ContainsKey(resid3))
            {
                Methods.SumValuta[resid3] += double.Parse(res3Sum);
            }
            else
            {
                Methods.SumValuta.Add(resid3, double.Parse(res3Sum));
            }




            /*Продолжить итоговая строка содержащая валюты и суммы по ним*/
            string itogo = "";
            string nowString = "";
            foreach(KeyValuePair<string,double> elem in Methods.SumValuta)
            {
                if(elem.Key == "") { continue; }
                Debug.WriteLine(elem.Key);
                nowString = String.Format("{0}: {1}, в рублях: {2}", elem.Key, elem.Value.ToString(), (elem.Value*Methods.courseVal[elem.Key]).ToString());
                itogo += nowString + "^l";     //elem.Key + ":    " + elem.Value.ToString() + "^l";
            }

            string[] ReplaceWithRef = 
            {
                val_1, price_1_Str, duration_1, sum_1,
                val_2, price_2_Str, duration_2, sum_2,
                val_3, price_3_Str, duration_3, sum_3,
                priceDayStr, durationDayStr, sumDayStr,
                valTr, s_sumTr,
                VairHair, s_SairHair,
                valOrg, s_sumOrg,
                valViza, s_sumViza,
                valMed, s_sumMed,
                resid1, res1Cost, res1Day, res1Sum,
                resid2, res2Cost, res2Day, res2Sum,
                resid3, res3Cost, res3Day, res3Sum,
                itogo
            };

            /******************************** массив для справки **********************************/
            string[] ReplaceWith =
            {
                number, data, fio, department_postAll, country1, city1, start1, end1, duration1,
                country2, city2, start2, end2, duration2, country3, city3, start3, end3, duration3,
                durationTotal, org, reason, numDocReason, dateDocReason, costMission, finSource,
                supperFin, bonus1_bonus2, purpose, mission, roadMap, competitionCash, extra, itogo
            };

            

            /*************************** формирование документов ************************/

            /*Создание записки*/
            Methods.replaceNote(FindText, ReplaceWith, @"\reference.docx");

            /*Создание справки*/
            Methods.replaceNote(FindTextRef, ReplaceWithRef, @"\spravka.docx");

        }
    }
}
