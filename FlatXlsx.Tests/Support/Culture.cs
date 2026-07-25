using System.Globalization;

namespace FlatXlsx.Tests.Support;

/// <summary>
/// Runs a piece of work under a chosen culture on its own thread.
/// </summary>
/// <remarks>
/// Culture is ambient state, and tests run in parallel; setting it in place would let one test
/// decide another test's outcome. A dedicated thread keeps the setting where it belongs.
/// </remarks>
static class Culture
{
    /// <param name="formatting">Decides how values are formatted. Pass null to leave it alone.</param>
    /// <param name="messages">Decides which translation of a message is used.</param>
    public static T Under<T>(string? formatting, string? messages, Func<T> body)
    {
        T result = default!;
        Exception? failure = null;

        var thread = new Thread(() =>
        {
            try
            {
                if (formatting != null)
                    CultureInfo.CurrentCulture = new CultureInfo(formatting);
                if (messages != null)
                    CultureInfo.CurrentUICulture = new CultureInfo(messages);
                result = body();
            }
            catch (Exception ex)
            {
                failure = ex;
            }
        });

        thread.Start();
        thread.Join();

        if (failure != null)
            throw new InvalidOperationException($"failed under formatting={formatting}, messages={messages}", failure);
        return result;
    }

    /// <summary>Captures the message of the exception the work is expected to throw.</summary>
    public static string MessageFrom(string messages, Action body) => Under(null, messages, () =>
    {
        try
        {
            body();
        }
        catch (Exception ex)
        {
            return ex.Message;
        }
        throw new InvalidOperationException("the work was expected to fail, but it succeeded");
    });
}
