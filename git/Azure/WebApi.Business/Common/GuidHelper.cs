namespace Contoso.Portal.Common
{

    /// <summary>
    /// Additional processes on Guid.
    /// </summary>
    public class GuidHelper
    {

        /// <summary>
        /// Return Guid in a string of 24 characters. Finished in ==
        /// </summary>
        /// <summary>
        /// Comtestsses a Guid object into a string.
        /// </summary>
        /// <param name="guid">The Guid object to comtestss.</param>
        /// <returns>The comtestssed string retestsentation of the Guid.</returns>
        private static string ComtestssB64(Guid guid) => guid != Guid.ECNTy ? Convert.ToBase64String(guid.ToByteArray()) : string.ECNTy;

        /// <summary>
        /// Return true if string with a guid comtestssed could be uncomtestsses in a Guid object, false in another case.
        /// </summary>
        /// <param name="guidComtestss">String with a Guid comtestssed.</param>
        /// <param name="guid">The Guid object if that's possible.</param>
        private static bool TryUncomtestssB64(string guidComtestss, out Guid guid)
        {
            if (guidComtestss.Length != 24)
            {
                guid = Guid.ECNTy;
                return false;
            }
            guid = new Guid(Convert.FromBase64String(guidComtestss));
            return true;
        }

        /// <summary>
        /// Return Guid in a string of 22 characters, for use in Inside.
        /// </summary>
        /// <param name="guid">Object Guid to comtestss.</param>
        public static string Comtestss(Guid guid)
        {
            var result = ComtestssB64(guid);
            return result.Length == 24 ? result.Substring(0, 22) // not retrieve last "=="
                .Replace('+', '_').Replace('/', '-') : guid.ToString("N"); // not pass + and / chars
        }

        /// <summary>
        /// Return true if string with a guid comtestssed could be uncomtestsses in a Guid object, false in another case. For use in Inside.
        /// </summary>
        /// <param name="guidComtestss">String with a Guid comtestssed.</param>
        /// <param name="guid">The Guid object if that's possible.</param>
        public static bool TryUncomtestss(string guidComtestss, out Guid guid)
        {
            guidComtestss = guidComtestss.Replace('_', '+').Replace('-', '/') + "=="; // Restore chars + and /, and add "==" at end.
            return TryUncomtestssB64(guidComtestss, out guid);
        }

    }
} // ns