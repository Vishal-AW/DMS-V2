import { WebPartContext } from "@microsoft/sp-webpart-base";
import { CreateItem } from "../DAL/Commonfile";

function createSendMailItem(context: WebPartContext, savedata: any) {
    return CreateItem(context.pageContext.web.absoluteUrl, context.spHttpClient, "DMS_SendEmail", savedata);
}

export async function TileSendMail(context: WebPartContext, docinfo: any) {
    let MailBody = "Dear User,";
    MailBody += "<br><br>I hope this message finds you well.<br><br>";

    if (docinfo.Status === "PendingWithPublisher") {
        MailBody += "I kindly request your approval for the document, <b>" + docinfo.DocName + "</b>. ";
        MailBody += "Please review the content at your convenience and let us know if any revisions or adjustments are required.";
    }
    else if (docinfo.Status === "PendingWithPM" || docinfo.Status === "Published") {
        MailBody += "We are pleased to inform you that the document titled <b>" + docinfo.DocName + "</b> has been reviewed and published.";
    }
    else if (docinfo.Status === "Rejected") {
        MailBody += "After reviewing the document titled <b>" + docinfo.DocName + "</b>, we regret to inform you that it has been <b>Rejected</b>.";
    }

    MailBody += "<br><br><b>Document Details:</b>";
    MailBody += "<br><b>Document Name:</b> " + docinfo.DocName;
    MailBody += "<br><b>Uploaded By:</b> " + docinfo.AuthorTitle;
    MailBody += "<br><b>Tile Name:</b> " + docinfo.TileName;
    MailBody += "<br><b>Document Path:</b> " + docinfo.FolderPath;

    // const actionBy = (docinfo.Status === "Rejected") ? "Rejected By" : "Approved By";
    // MailBody += `<br><b>${actionBy}:</b> ${context.pageContext.user.displayName}`;


    if (docinfo.Status != "PendingWithPM") {
        const actionBy = (docinfo.Status === "Rejected") ? "Rejected By" : "Approved By";
        MailBody += `<br><b>${actionBy}:</b> ${context.pageContext.user.displayName}`;
    }


    MailBody += "<br>";

    const subject = docinfo.Sub;
    const Mail = {
        From: context.pageContext.user.email,
        Subject: subject,
        Body: MailBody,
        To: docinfo.To,
        FolderPath: docinfo.FolderPath,
        DocName: docinfo.DocName,
        LID: docinfo.ID,
        LibraryName: docinfo.libraryName,
        Status: docinfo.Status,

    };

    return await createSendMailItem(context, Mail);
}
