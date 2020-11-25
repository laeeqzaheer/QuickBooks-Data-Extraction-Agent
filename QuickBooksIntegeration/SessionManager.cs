using QBFC13Lib;
using System;
using System.IO;

/// <summary>
/// Defines the <see cref="SessionManager" />.
/// </summary>
public sealed class SessionManager
{
    /// <summary>
    /// Defines the instance.
    /// </summary>
    private static readonly SessionManager instance = new SessionManager();

    /// <summary>
    /// Defines the isConnectionOpen.
    /// </summary>
    internal bool isConnectionOpen = false;

    /// <summary>
    /// Defines the isSessionBegun.
    /// </summary>
    internal bool isSessionBegun = false;

    /// <summary>
    /// Gets or sets a value indicating whether IsSessionBegun.
    /// </summary>
    public bool IsSessionBegun
    {
        get
        {
            return isSessionBegun;
        }
        set
        {
            isSessionBegun = value;
        }
    }

    /// <summary>
    /// Gets or sets a value indicating whether IsConnectionOpen.
    /// </summary>
    public bool IsConnectionOpen
    {
        get
        {
            return isConnectionOpen;
        }
        set
        {
            isConnectionOpen = value;
        }
    }

    /// <summary>
    /// Defines the session.
    /// </summary>
    internal QBSessionManager session = new QBSessionManager();

    /// <summary>
    /// Initializes static members of the <see cref="SessionManager"/> class.
    /// </summary>
    static SessionManager()
    {
    }

    /// <summary>
    /// Prevents a default instance of the <see cref="SessionManager"/> class from being created.
    /// </summary>
    private SessionManager()
    {
    }

    /// <summary>
    /// Gets the Instance.
    /// </summary>
    public static SessionManager Instance
    {
        get
        {
            return instance;
        }
    }

    /// <summary>
    /// The OpenConnection.
    /// </summary>
    /// <param name="appName">The appName<see cref="string"/>.</param>
    /// <param name="logsFile">The logsFile<see cref="StreamWriter"/>.</param>
    public void OpenConnection(string appName, StreamWriter logsFile)
    {
        try
        {
            if (!isConnectionOpen)
            {
                // Args: App ID, App Name
                isConnectionOpen = true;
                session.OpenConnection("", appName);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Warning Error Occured During Execution, Please Refer to Logs...\n");
            logsFile.WriteLine(DateTime.Now + "\tERROR\tError Message:\t" + ex.Message);
            logsFile.WriteLine();
            logsFile.WriteLine(DateTime.Now + "\tERROR\tError Message:\tStack Trace:\t" + ex.StackTrace);
            logsFile.WriteLine();
            //isConnectionOpen = false;
            CloseConnection(logsFile);
            //Program.Instance.closeQuickBooks(logsFile);

        }
    }

    /// <summary>
    /// The BeginSession.
    /// </summary>
    /// <param name="filePath">The filePath<see cref="string"/>.</param>
    /// <param name="logsFile">The logsFile<see cref="StreamWriter"/>.</param>
    public void BeginSession(string filePath, StreamWriter logsFile)
    {
        try
        {
            if (!isSessionBegun)
            {
                // By passing `""` for QB Filename, begin session with currently open QB Filename
                // Args: QB Filename, ENOpenMode
                //                session.BeginSession("C:\\QuickBooksSetup\\Sample Stores\\Popeye-Newport Plaza's New File.QBW", ENOpenMode.omSingleUser);
                isSessionBegun = true;
                session.BeginSession(filePath, ENOpenMode.omMultiUser);

            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Warning Error Occured During Execution, Please Refer to Logs...\n");
            logsFile.WriteLine(DateTime.Now + "\tERROR\tError Message:\t" + ex.Message);
            logsFile.WriteLine();
            logsFile.WriteLine(DateTime.Now + "\tERROR\tError Message:\tStack Trace:\t" + ex.StackTrace);
            logsFile.WriteLine();
            //isSessionBegun = false;
            EndSession(logsFile);
            CloseConnection(logsFile);

        }
    }

    /// <summary>
    /// The EndSession.
    /// </summary>
    /// <param name="logsFile">The logsFile<see cref="StreamWriter"/>.</param>
    public void EndSession(StreamWriter logsFile)
    {
        try
        {
            if (isSessionBegun)
            {
                session.EndSession();
                isSessionBegun = false;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Warning Error Occured During Execution, Please Refer to Logs...\n");
            logsFile.WriteLine(DateTime.Now + "\tERROR\tError Message:\t" + ex.Message);
            logsFile.WriteLine();
            logsFile.WriteLine(DateTime.Now + "\tERROR\tError Message:\tStack Trace:\t" + ex.StackTrace);
            logsFile.WriteLine();
        }
    }

    /// <summary>
    /// The CloseConnection.
    /// </summary>
    /// <param name="logsFile">The logsFile<see cref="StreamWriter"/>.</param>
    public void CloseConnection(StreamWriter logsFile)
    {
        try
        {
            if (isConnectionOpen)
            {
                session.CloseConnection();
                isConnectionOpen = false;
            }

        }
        catch (Exception ex)
        {
            Console.WriteLine("Warning Error Occured During Execution, Please Refer to Logs...\n");
            logsFile.WriteLine(DateTime.Now + "\tERROR\tError Message:\t" + ex.Message);
            logsFile.WriteLine();
            logsFile.WriteLine(DateTime.Now + "\tERROR\tError Message:\tStack Trace:\t" + ex.StackTrace);
            logsFile.WriteLine();
        }
    }

    /// <summary>
    /// The GetCurrentCompanyFileName.
    /// </summary>
    /// <returns>The <see cref="string"/>.</returns>
    public string GetCurrentCompanyFileName()
    {
        return session.GetCurrentCompanyFileName();
    }

    /// <summary>
    /// The CreateMsgSetRequest.
    /// </summary>
    /// <returns>The <see cref="IMsgSetRequest"/>.</returns>
    public IMsgSetRequest CreateMsgSetRequest()
    {
        // Args: Request Country, QBFC Major Version, QBFC Minor Version
        return session.CreateMsgSetRequest("US", 13, 0);
    }

    /// <summary>
    /// The DoRequests.
    /// </summary>
    /// <param name="requestSet">The requestSet<see cref="IMsgSetRequest"/>.</param>
    /// <param name="logsFile">The logsFile<see cref="StreamWriter"/>.</param>
    /// <returns>The <see cref="IMsgSetResponse"/>.</returns>
    public IMsgSetResponse DoRequests(IMsgSetRequest requestSet, StreamWriter logsFile)
    {
        try { return session.DoRequests(requestSet); }
        catch (Exception ex)
        {
            isSessionBegun = false;
            Console.WriteLine("Warning Error Occured During Execution, Please Refer to Logs...\n");
            logsFile.WriteLine(DateTime.Now + "\tERROR\tError Message:\t" + ex.Message);
            logsFile.WriteLine();
            logsFile.WriteLine(DateTime.Now + "\tERROR\tError Message:\tStack Trace:\t" + ex.StackTrace);
            logsFile.WriteLine();
            return null;
        }
    }
}
