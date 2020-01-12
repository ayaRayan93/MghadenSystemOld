﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using DevExpress.XtraTab;
using DevExpress.XtraGrid;
using DevExpress.XtraTab.ViewInfo;
using DevExpress.XtraNavBar;
using MySql.Data.MySqlClient;

namespace MainSystem
{
    class ExpensesSystem
    {
    }
    public partial class MainForm
    {
        public static XtraTabControl tabControlExpenses;

        public void initializeBranch()
        {
            //bankSystem
            tabControlExpenses = xtraTabControlExpenses;
        }

        private void navBarItemMainSubReport_LinkClicked(object sender, NavBarLinkEventArgs e)
        {
            try
            {
                restForeColorOfNavBarItem();
                NavBarItem navBarItem = (NavBarItem)sender;
                navBarItem.Appearance.ForeColor = Color.Blue;

                if (!xtraTabControlExpenses.Visible)
                    xtraTabControlExpenses.Visible = true;

                XtraTabPage xtraTabPage = getTabPage(xtraTabControlExpenses, "تكويد المصروفات");
                if (xtraTabPage == null)
                {
                    xtraTabControlExpenses.TabPages.Add("تكويد المصروفات");
                    xtraTabPage = getTabPage(xtraTabControlExpenses, "تكويد المصروفات");
                }
                xtraTabPage.Controls.Clear();

                xtraTabControlExpenses.SelectedTabPage = xtraTabPage;
                Main_Sub objFormExpenses = new Main_Sub(xtraTabControlExpenses);
                objFormExpenses.TopLevel = false;

                xtraTabPage.Controls.Add(objFormExpenses);
                objFormExpenses.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
                objFormExpenses.Dock = DockStyle.Fill;
                objFormExpenses.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void navBarItemExpenseRecord_LinkClicked(object sender, NavBarLinkEventArgs e)
        {
            try
            {
                restForeColorOfNavBarItem();
                NavBarItem navBarItem = (NavBarItem)sender;
                navBarItem.Appearance.ForeColor = Color.Blue;

                if (!xtraTabControlExpenses.Visible)
                    xtraTabControlExpenses.Visible = true;

                XtraTabPage xtraTabPage = getTabPage(xtraTabControlExpenses, "تسجيل مصروف");
                if (xtraTabPage == null)
                {
                    xtraTabControlExpenses.TabPages.Add("تسجيل مصروف");
                    xtraTabPage = getTabPage(xtraTabControlExpenses, "تسجيل مصروف");
                }
                xtraTabPage.Controls.Clear();

                xtraTabControlExpenses.SelectedTabPage = xtraTabPage;
                SafeExpense_Record2 objFormExpenses = new SafeExpense_Record2(xtraTabControlExpenses);
                objFormExpenses.TopLevel = false;

                xtraTabPage.Controls.Add(objFormExpenses);
                objFormExpenses.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
                objFormExpenses.Dock = DockStyle.Fill;
                objFormExpenses.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void navBarItemSafeExpenseIncomeRecord_LinkClicked(object sender, NavBarLinkEventArgs e)
        {
            try
            {
                restForeColorOfNavBarItem();
                NavBarItem navBarItem = (NavBarItem)sender;
                navBarItem.Appearance.ForeColor = Color.Blue;

                if (!xtraTabControlExpenses.Visible)
                    xtraTabControlExpenses.Visible = true;

                XtraTabPage xtraTabPage = getTabPage(xtraTabControlExpenses, "تسجيل وارد لمصروف");
                if (xtraTabPage == null)
                {
                    xtraTabControlExpenses.TabPages.Add("تسجيل وارد لمصروف");
                    xtraTabPage = getTabPage(xtraTabControlExpenses, "تسجيل وارد لمصروف");
                }
                xtraTabPage.Controls.Clear();

                xtraTabControlExpenses.SelectedTabPage = xtraTabPage;
                SafeExpenseIncome_Record objFormExpenses = new SafeExpenseIncome_Record(xtraTabControlExpenses);
                objFormExpenses.TopLevel = false;

                xtraTabPage.Controls.Add(objFormExpenses);
                objFormExpenses.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
                objFormExpenses.Dock = DockStyle.Fill;
                objFormExpenses.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void navBarItemExpensesTransitionsReport_LinkClicked(object sender, NavBarLinkEventArgs e)
        {
            try
            {
                restForeColorOfNavBarItem();
                NavBarItem navBarItem = (NavBarItem)sender;
                navBarItem.Appearance.ForeColor = Color.Blue;

                if (!xtraTabControlExpenses.Visible)
                    xtraTabControlExpenses.Visible = true;

                XtraTabPage xtraTabPage = getTabPage(xtraTabControlExpenses, "حركة المصروفات بالخزنة");
                if (xtraTabPage == null)
                {
                    xtraTabControlExpenses.TabPages.Add("حركة المصروفات بالخزنة");
                    xtraTabPage = getTabPage(xtraTabControlExpenses, "حركة المصروفات بالخزنة");
                }
                xtraTabPage.Controls.Clear();

                xtraTabControlExpenses.SelectedTabPage = xtraTabPage;
                Expenses_Transitions_Report objFormExpenses = new Expenses_Transitions_Report(xtraTabControlExpenses, this);
                objFormExpenses.TopLevel = false;

                xtraTabPage.Controls.Add(objFormExpenses);
                objFormExpenses.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
                objFormExpenses.Dock = DockStyle.Fill;
                objFormExpenses.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void navBarItemSubExpensesTransitionsReport_LinkClicked(object sender, NavBarLinkEventArgs e)
        {
            try
            {
                restForeColorOfNavBarItem();
                NavBarItem navBarItem = (NavBarItem)sender;
                navBarItem.Appearance.ForeColor = Color.Blue;

                if (!xtraTabControlExpenses.Visible)
                    xtraTabControlExpenses.Visible = true;

                XtraTabPage xtraTabPage = getTabPage(xtraTabControlExpenses, "حركة المصروفات بالمصروف");
                if (xtraTabPage == null)
                {
                    xtraTabControlExpenses.TabPages.Add("حركة المصروفات بالمصروف");
                    xtraTabPage = getTabPage(xtraTabControlExpenses, "حركة المصروفات بالمصروف");
                }
                xtraTabPage.Controls.Clear();

                xtraTabControlExpenses.SelectedTabPage = xtraTabPage;
                SubExpensesTransitions_Report objFormExpenses = new SubExpensesTransitions_Report(xtraTabControlExpenses);
                objFormExpenses.TopLevel = false;

                xtraTabPage.Controls.Add(objFormExpenses);
                objFormExpenses.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
                objFormExpenses.Dock = DockStyle.Fill;
                objFormExpenses.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void bindUpdateExpenseForm(DataRowView rows, Expenses_Transitions_Report ExpensesTransitionsReport)
        {
            if (!xtraTabControlExpenses.Visible)
                xtraTabControlExpenses.Visible = true;

            XtraTabPage xtraTabPage = getTabPage(xtraTabControlExpenses, "تعديل مصروف");

            if (xtraTabPage == null)
            {
                xtraTabControlExpenses.TabPages.Add("تعديل مصروف");
                xtraTabPage = getTabPage(xtraTabControlExpenses, "تعديل مصروف");
            }
            xtraTabPage.Controls.Clear();

            xtraTabControlExpenses.SelectedTabPage = xtraTabPage;

            SafeExpense_Update objForm = new SafeExpense_Update(rows, ExpensesTransitionsReport, xtraTabControlExpenses, this);
            objForm.TopLevel = false;

            xtraTabPage.Controls.Add(objForm);
            objForm.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            objForm.Dock = DockStyle.Fill;
            objForm.Show();
        }

        public void bindUpdateIncomeExpenseForm(DataRowView rows, Expenses_Transitions_Report ExpensesTransitionsReport)
        {
            if (!xtraTabControlExpenses.Visible)
                xtraTabControlExpenses.Visible = true;

            XtraTabPage xtraTabPage = getTabPage(xtraTabControlExpenses, "تعديل ايداع مصروف");

            if (xtraTabPage == null)
            {
                xtraTabControlExpenses.TabPages.Add("تعديل ايداع مصروف");
                xtraTabPage = getTabPage(xtraTabControlExpenses, "تعديل ايداع مصروف");
            }
            xtraTabPage.Controls.Clear();

            xtraTabControlExpenses.SelectedTabPage = xtraTabPage;

            SafeExpenseIncome_Update objForm = new SafeExpenseIncome_Update(rows, ExpensesTransitionsReport, xtraTabControlExpenses, this);
            objForm.TopLevel = false;

            xtraTabPage.Controls.Add(objForm);
            objForm.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            objForm.Dock = DockStyle.Fill;
            objForm.Show();
        }
    }
}
