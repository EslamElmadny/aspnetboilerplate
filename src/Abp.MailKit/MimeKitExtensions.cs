//
// Author: Jeffrey Stedfast <jestedfa@microsoft.com>
//
// Copyright (c) 2013-2017 Xamarin Inc. (www.xamarin.com)
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE.
//

using MimeKit;
using MimeKit.IO;
using System;
using System.Collections.Generic;
using System.Net.Mail;
using System.Text;

namespace Abp.MailKit
{
    public static class EmailExtensions
    {
        public static MimeMessage ToMimeMessage(this MailMessage message)
        {
            if (message == null)
                throw new ArgumentNullException(nameof(message));

            var headers = new List<Header>();
            foreach (var field in message.Headers.AllKeys)
            {
                foreach (var value in message.Headers.GetValues(field))
                    headers.Add(new Header(field, value));
            }

            var msg = new MimeMessage(ParserOptions.Default, headers, RfcComplianceMode.Strict);
            MimeEntity body = null;

            // Note: If the user has already sent their MailMessage via System.Net.Mail.SmtpClient,
            // then the following MailMessage properties will have been merged into the Headers, so
            // check to make sure our MimeMessage properties are empty before adding them.
            if (message.Sender != null)
                msg.Sender = message.Sender.ToMailboxAddress();

            if (message.From != null)
            {
                msg.Headers.Replace(HeaderId.From, string.Empty);
                msg.From.Add(message.From.ToMailboxAddress());
            }

            if (message.ReplyToList.Count > 0)
            {
                msg.Headers.Replace(HeaderId.ReplyTo, string.Empty);
                msg.ReplyTo.AddRange(message.ReplyToList.ToInternetAddressList());
            }

            if (message.To.Count > 0)
            {
                msg.Headers.Replace(HeaderId.To, string.Empty);
                msg.To.AddRange(message.To.ToInternetAddressList());
            }

            if (message.CC.Count > 0)
            {
                msg.Headers.Replace(HeaderId.Cc, string.Empty);
                msg.Cc.AddRange(message.CC.ToInternetAddressList());
            }

            if (message.Bcc.Count > 0)
            {
                msg.Headers.Replace(HeaderId.Bcc, string.Empty);
                msg.Bcc.AddRange(message.Bcc.ToInternetAddressList());
            }

            if (message.SubjectEncoding != null)
                msg.Headers.Replace(HeaderId.Subject, message.SubjectEncoding, message.Subject ?? string.Empty);
            else
                msg.Subject = message.Subject ?? string.Empty;

            switch (message.Priority)
            {
                case MailPriority.Normal:
                    msg.Headers.RemoveAll(HeaderId.XMSMailPriority);
                    msg.Headers.RemoveAll(HeaderId.Importance);
                    msg.Headers.RemoveAll(HeaderId.XPriority);
                    msg.Headers.RemoveAll(HeaderId.Priority);
                    break;
                case MailPriority.High:
                    msg.Headers.Replace(HeaderId.Priority, "urgent");
                    msg.Headers.Replace(HeaderId.Importance, "high");
                    msg.Headers.Replace(HeaderId.XPriority, "2 (High)");
                    break;
                case MailPriority.Low:
                    msg.Headers.Replace(HeaderId.Priority, "non-urgent");
                    msg.Headers.Replace(HeaderId.Importance, "low");
                    msg.Headers.Replace(HeaderId.XPriority, "4 (Low)");
                    break;
            }

            if (!string.IsNullOrEmpty(message.Body))
            {
                var text = new TextPart(message.IsBodyHtml ? "html" : "plain");
                text.SetText(message.BodyEncoding ?? Encoding.UTF8, message.Body);
                body = text;
            }

            if (message.AlternateViews.Count > 0)
            {
                var alternative = new MultipartAlternative();

                if (body != null)
                    alternative.Add(body);

                foreach (var view in message.AlternateViews)
                {
                    var part = GetMimePart(view);

                    if (view.BaseUri != null)
                        part.ContentLocation = view.BaseUri;

                    if (view.LinkedResources.Count > 0)
                    {
                        var type = part.ContentType.MediaType + "/" + part.ContentType.MediaSubtype;
                        var related = new MultipartRelated();

                        related.ContentType.Parameters.Add("type", type);

                        if (view.BaseUri != null)
                            related.ContentLocation = view.BaseUri;

                        related.Add(part);

                        foreach (var resource in view.LinkedResources)
                        {
                            part = GetMimePart(resource);

                            if (resource.ContentLink != null)
                                part.ContentLocation = resource.ContentLink;

                            related.Add(part);
                        }

                        alternative.Add(related);
                    }
                    else
                    {
                        alternative.Add(part);
                    }
                }

                body = alternative;
            }

            if (body == null)
                body = new TextPart(message.IsBodyHtml ? "html" : "plain");

            if (message.Attachments.Count > 0)
            {
                var mixed = new Multipart("mixed");

                if (body != null)
                    mixed.Add(body);

                foreach (var attachment in message.Attachments)
                    mixed.Add(GetMimePart(attachment));

                body = mixed;
            }

            msg.Body = body;

            return msg;
        }

        private static MimePart GetMimePart(AttachmentBase item)
        {
            var mimeType = item.ContentType.ToString();
            var contentType = ContentType.Parse(mimeType);
            var attachment = item as Attachment;
            MimePart part;

            // if (contentType.MediaType.Equals("text", StringComparison.OrdinalIgnoreCase))
            //     part = new TextPart(contentType);
            // else
                part = new MimePart(contentType);

            if (attachment != null)
            {
                var disposition = attachment.ContentDisposition.ToString();
                part.ContentDisposition = ContentDisposition.Parse(disposition);
            }

            switch (item.TransferEncoding)
            {
                case System.Net.Mime.TransferEncoding.QuotedPrintable:
                    part.ContentTransferEncoding = ContentEncoding.QuotedPrintable;
                    break;
                case System.Net.Mime.TransferEncoding.Base64:
                    part.ContentTransferEncoding = ContentEncoding.Base64;
                    break;
                case System.Net.Mime.TransferEncoding.SevenBit:
                    part.ContentTransferEncoding = ContentEncoding.SevenBit;
                    break;
                    //case System.Net.Mime.TransferEncoding.EightBit:
                    //	part.ContentTransferEncoding = ContentEncoding.EightBit;
                    //	break;
            }

            if (item.ContentId != null)
                part.ContentId = item.ContentId;

            var stream = new MemoryBlockStream();
            item.ContentStream.CopyTo(stream);
            stream.Position = 0;

            part.ContentObject = new ContentObject(stream);

            return part;
        }

        private static InternetAddressList ToInternetAddressList(this MailAddressCollection addresses)
        {
            if (addresses == null)
                return null;

            var list = new InternetAddressList();
            foreach (var address in addresses)
                list.Add(address.ToMailboxAddress());

            return list;
        }

        private static MailboxAddress ToMailboxAddress(this MailAddress address)
        {
            return address != null ? new MailboxAddress(address.DisplayName, address.Address) : null;
        }
    }
}
