﻿/*	
 * 	
 *  This file is part of libfintx.
 *  
 *  Copyright (c) 2016 - 2020 Torsten Klinger
 * 	E-Mail: torsten.klinger@googlemail.com
 * 	
 * 	libfintx is free software; you can redistribute it and/or
 *	modify it under the terms of the GNU Lesser General Public
 * 	License as published by the Free Software Foundation; either
 * 	version 2.1 of the License, or (at your option) any later version.
 *	
 * 	libfintx is distributed in the hope that it will be useful,
 * 	but WITHOUT ANY WARRANTY; without even the implied warranty of
 * 	MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
 * 	Lesser General Public License for more details.
 *	
 * 	You should have received a copy of the GNU Lesser General Public
 * 	License along with libfintx; if not, write to the Free Software
 * 	Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA
 * 	
 */

using System;
using System.Threading.Tasks;

namespace libfintx
{
    public static class HKCCS
    {
        /// <summary>
        /// Transfer
        /// </summary>
        public static async Task<String> Init_HKCCS(FinTsClient client, string ReceiverName, string ReceiverIBAN, string ReceiverBIC, decimal Amount, string Usage)
        {
            Log.Write("Starting job HKCCS: Transfer money");
            var connectionDetails = client.ConnectionDetails;
            string segments = string.Empty;

            string sepaMessage = string.Empty;

            client.SEGNUM = SEGNUM.SETInt(3);

            if (client.HISPAS == 1)
            {
                segments = "HKCCS:" + client.SEGNUM + ":1+" + connectionDetails.Iban + ":" + connectionDetails.Bic + "+urn?:iso?:std?:iso?:20022?:tech?:xsd?:pain.001.001.03+@@";
                sepaMessage = pain00100103.Create(connectionDetails.AccountHolder, connectionDetails.Iban, connectionDetails.Bic, ReceiverName, ReceiverIBAN, ReceiverBIC, Amount, Usage, new DateTime(1999, 1, 1));
            }
            else if (client.HISPAS == 2)
            {
                segments = "HKCCS:" + client.SEGNUM + ":1+" + connectionDetails.Iban + ":" + connectionDetails.Bic + "+urn?:iso?:std?:iso?:20022?:tech?:xsd?:pain.001.002.03+@@";
                sepaMessage = pain00100203.Create(connectionDetails.AccountHolder, connectionDetails.Iban, connectionDetails.Bic, ReceiverName, ReceiverIBAN, ReceiverBIC, Amount, Usage, new DateTime(1999, 1, 1));
            }
            else if (client.HISPAS == 3)
            {
                segments = "HKCCS:" + client.SEGNUM + ":1+" + connectionDetails.Iban + ":" + connectionDetails.Bic + "+urn?:iso?:std?:iso?:20022?:tech?:xsd?:pain.001.003.03+@@";
                sepaMessage = pain00100303.Create(connectionDetails.AccountHolder, connectionDetails.Iban, connectionDetails.Bic, ReceiverName, ReceiverIBAN, ReceiverBIC, Amount, Usage, new DateTime(1999, 1, 1));
            }

            segments = segments.Replace("@@", "@" + (sepaMessage.Length - 1) + "@") + sepaMessage;

            if (Helper.IsTANRequired("HKCCS"))
            {
                client.SEGNUM = SEGNUM.SETInt(4);
                segments = HKTAN.Init_HKTAN(client, segments, "HKCCS");
            }

            var message = FinTSMessage.Create(client, client.HNHBS, client.HNHBK, segments, client.HIRMS);
            var response = await FinTSMessage.Send(client, message);

            Helper.Parse_Message(client, response);

            return response;
        }
    }
}
