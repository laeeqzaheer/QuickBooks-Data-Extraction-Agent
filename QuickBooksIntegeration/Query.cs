using QBFC13Lib;
using System;
using System.Collections.Generic;
using System.IO;

/// <summary>
/// Defines the <see cref="Query" />.
/// </summary>
public class Query
{
    /// <summary>
    /// The GetQBAccounts.
    /// </summary>
    /// <param name="logsFile">The logsFile<see cref="StreamWriter"/>.</param>
    /// <returns>The <see cref="List{dynamic}"/>.</returns>
    public List<dynamic> GetQBAccounts(StreamWriter logsFile)
    {

        IMsgSetRequest requestSet = SessionManager.Instance.CreateMsgSetRequest();
        requestSet.Attributes.OnError = ENRqOnError.roeStop;
        requestSet.AppendAccountQueryRq();
        IMsgSetResponse responeSet = SessionManager.Instance.DoRequests(requestSet, logsFile);
        IResponseList responseList = responeSet.ResponseList;
        List<dynamic> accountsArray = new List<dynamic>();
        var account_number = "";
        try
        {
            for (int i = 0; i < responseList.Count; i++)
            {
                IResponse response = responseList.GetAt(i);

                if (response.StatusCode == 0)
                {
                    IAccountRetList AccountRet = (IAccountRetList)response.Detail;
                    Console.WriteLine("Number Of Records Being Fetched Are:\t" + AccountRet.Count + "\n");
                    for (int j = 0; j < AccountRet.Count; j++)
                    {
                        IAccountRet account = AccountRet.GetAt(j);
                        if (account.AccountNumber != null)
                        {
                            account_number = account.AccountNumber.GetValue();
                        }
                        else
                        {
                            account_number = "Account Number Not Defined...";
                        }
                        accountsArray.Add(new string[] { account.Name.GetValue().ToString(), account_number, account.AccountType.GetValue().ToString(), account.TotalBalance.GetValue().ToString() });
                    }

                }

            }
        }
        catch (Exception ex)
        {
            logsFile.WriteLine(DateTime.Now + "\tERROR\tError Message:\t" + ex.Message);
            logsFile.WriteLine();
            logsFile.WriteLine(DateTime.Now + "\tERROR\tError Message:\tStack Trace:\t" + ex.StackTrace);
            logsFile.WriteLine();
        }

        return accountsArray;
    }

    /// <summary>
    /// The GetVendorsData.
    /// </summary>
    /// <param name="logsFile">The logsFile<see cref="StreamWriter"/>.</param>
    /// <returns>The <see cref="List{dynamic}"/>.</returns>
    public List<dynamic> GetVendorsData(StreamWriter logsFile)
    {
        IMsgSetRequest requestSet = SessionManager.Instance.CreateMsgSetRequest();
        requestSet.Attributes.OnError = ENRqOnError.roeStop;
        IVendorQuery VendorQueryRq = requestSet.AppendVendorQueryRq();
        IMsgSetResponse responeSet = SessionManager.Instance.DoRequests(requestSet, logsFile);
        IResponseList responseList = responeSet.ResponseList;
        List<dynamic> vendorsArray = new List<dynamic>();
        var vendor_name = "";
        var vendor_balance = "";
        var account_number = "";
        try
        {
            for (int i = 0; i < responseList.Count; i++)
            {
                IResponse response = responseList.GetAt(i);
                if (response.StatusCode == 0)
                {
                    IVendorRetList vendorsList = (IVendorRetList)response.Detail;
                    Console.WriteLine("Number Of Records Being Fetched Are:\t" + vendorsList.Count + "\n");
                    for (int j = 0; j < vendorsList.Count; j++)
                    {
                        IVendorRet vendor = vendorsList.GetAt(j);
                        if (vendor.Name != null)
                        {
                            vendor_name = vendor.Name.GetValue();
                        }
                        else
                        {
                            vendor_name = "Name Not Defined...";
                        }
                        if (vendor.Balance != null)
                        {
                            vendor_balance = vendor.Balance.GetValue().ToString();
                        }
                        else
                        {
                            vendor_balance = "Balance Not Defined...";
                        }
                        if (vendor.AccountNumber != null)
                        {
                            account_number = vendor.AccountNumber.GetValue();
                        }
                        else
                        {
                            account_number = "Account Number Not Defined...";
                        }
                        vendorsArray.Add(new string[] { vendor_name, account_number, vendor_balance });
                    }

                }
            }
        }
        catch (Exception ex)
        {
            logsFile.WriteLine(DateTime.Now + "\tERROR\tError Message:\t" + ex.Message);
            logsFile.WriteLine();
            logsFile.WriteLine(DateTime.Now + "\tERROR\tError Message:\tStack Trace:\t" + ex.StackTrace);
            logsFile.WriteLine();
        }

        return vendorsArray;
    }

    /// <summary>
    /// The GetCustomersData.
    /// </summary>
    /// <param name="logsFile">The logsFile<see cref="StreamWriter"/>.</param>
    /// <returns>The <see cref="List{dynamic}"/>.</returns>
    public List<dynamic> GetCustomersData(StreamWriter logsFile)
    {
        IMsgSetRequest requestSet = SessionManager.Instance.CreateMsgSetRequest();
        requestSet.Attributes.OnError = ENRqOnError.roeStop;
        ICustomerQuery CustomerQueryRq = requestSet.AppendCustomerQueryRq();
        IMsgSetResponse responeSet = SessionManager.Instance.DoRequests(requestSet, logsFile);
        IResponseList responseList = responeSet.ResponseList;
        List<dynamic> customersArray = new List<dynamic>();
        var accountNumber = "";
        try
        {
            for (int i = 0; i < responseList.Count; i++)
            {
                IResponse response = responseList.GetAt(i);
                if (response.StatusCode == 0)
                {
                    ICustomerRetList customerList = (ICustomerRetList)response.Detail;
                    Console.WriteLine("Number Of Records Being Fetched Are:\t" + customerList.Count + "\n");
                    for (int j = 0; j < customerList.Count; j++)
                    {
                        ICustomerRet customer = customerList.GetAt(j);
                        if (customer.AccountNumber != null)
                        {
                            accountNumber = customer.AccountNumber.GetValue();
                        }
                        else
                        {
                            accountNumber = "Account Number Not Defined...";
                        }
                        customersArray.Add(new string[] { customer.Name.GetValue().ToString(), accountNumber, customer.TotalBalance.GetValue().ToString() });
                    }

                }
            }
        }

        catch (Exception ex)
        {
            logsFile.WriteLine(DateTime.Now + "\tERROR\tError Message:\t" + ex.Message);
            logsFile.WriteLine();
            logsFile.WriteLine(DateTime.Now + "\tERROR\tError Message:\tStack Trace:\t" + ex.StackTrace);
            logsFile.WriteLine();
        }

        return customersArray;
    }
}
