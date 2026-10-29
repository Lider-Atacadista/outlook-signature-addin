/**
 * Suplemento de Assinatura Automatica - Lider Atacadista & Bem Familia
 * Plataforma: Microsoft 365 (Novo Outlook, OWA e Outlook Classico)
 * Hospedagem: GitHub Pages (Lider-Atacadista/outlook-signature-addin)
 */

Office.onReady();

function onNewMessageComposeHandler(event) {
    try {
        var userProfile = Office.context.mailbox.userProfile;
        var email = (userProfile.emailAddress || "").toLowerCase().trim();
        var displayName = userProfile.displayName || "";
        
        // Determina a marca pelo dominio do UPN / E-mail
        var isBemFamilia = (email.indexOf("@bemfamilia.ind.br") !== -1);
        
        var catalogUrl = "https://lider-atacadista.github.io/outlook-signature-addin/users.json?v=" + new Date().getTime();

        if (typeof fetch !== "undefined") {
            fetch(catalogUrl)
                .then(function(response) {
                    if (!response.ok) {
                        throw new Error("HTTP " + response.status);
                    }
                    return response.json();
                })
                .then(function(usersCatalog) {
                    var userData = usersCatalog[email] || {};
                    var html = buildSignatureHtml(userData, isBemFamilia, email, displayName);
                    applySignature(html, event);
                })
                .catch(function(err) {
                    console.warn("Aviso ao carregar users.json, utilizando dados do perfil:", err);
                    var fallbackHtml = buildSignatureHtml({}, isBemFamilia, email, displayName);
                    applySignature(fallbackHtml, event);
                });
        } else {
            var fallbackHtml = buildSignatureHtml({}, isBemFamilia, email, displayName);
            applySignature(fallbackHtml, event);
        }
    } catch (ex) {
        console.error("Excecao no tratador de eventos:", ex);
        event.completed();
    }
}

function applySignature(htmlContent, event) {
    Office.context.mailbox.item.body.setSignatureAsync(
        htmlContent,
        { coercionType: Office.CoercionType.Html, asyncContext: event },
        function(asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error("Erro ao aplicar assinatura:", asyncResult.error.message);
            }
            // Obrigatorio sinalizar a conclusao do evento para o Outlook
            event.completed();
        }
    );
}

function sanitize(str) {
    if (!str) return "";
    return String(str)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/[\u007F-\uFFFF]/g, function(c) {
            return "&#" + c.charCodeAt(0) + ";";
        });
}

function buildSignatureHtml(userData, isBemFamilia, email, displayName) {
    var companyName = isBemFamilia ? "Bem Fam&iacute;lia" : "Lider Atacadista";
    
    // Imagens hospedadas diretamente no repositorio GitHub Pages
    var baseUrl = "https://lider-atacadista.github.io/outlook-signature-addin/images/";
    var logoUrl = isBemFamilia 
        ? baseUrl + "Logo-BemFamilia.png" 
        : baseUrl + "logo.png";
    var linkRede = isBemFamilia 
        ? "https://www.instagram.com/deliciasbemfamilia" 
        : "https://www.instagram.com/lider_atacadista";

    var icoMail = baseUrl + "ico-mail.png";
    var icoTel = baseUrl + "ico-tel.png";
    var icoCel = baseUrl + "ico-cel.png";

    var name = sanitize(userData.name || displayName || "Colaborador");
    var title = sanitize(userData.title || "");
    var department = sanitize(userData.department || "");
    var address = sanitize(userData.address || (isBemFamilia ? "Rod. SC 445, Km 06" : "Rua Jos\u00E9 Alberto Nunes, s/n - Humait\u00E1 de Cima"));
    var city = sanitize(userData.city || (isBemFamilia ? "I\u00E7ara - SC" : "Tubar\u00E3o - SC"));
    var phone = sanitize(userData.phone || "");
    var pager = userData.pager ? (" | Ramal: " + sanitize(userData.pager)) : "";
    var mobile = sanitize(userData.mobile || "");

    var cargoDepto = "";
    if (title && department) {
        cargoDepto = title + " | " + department;
    } else if (title) {
        cargoDepto = title;
    } else if (department) {
        cargoDepto = department;
    }

    var subLine = "<b>" + companyName + "</b>";
    if (address) subLine += " - <i>" + address + "</i>";
    if (city) subLine += " - <b>" + city + "</b>";

    var foneRow = "";
    if (phone) {
        foneRow = '<tr>' +
            '<td style="vertical-align: middle; padding-right: 8px; width: 20px;">' +
                '<img src="' + icoTel + '" width="16" height="16" alt="Telefone" style="border: 0; display: block;">' +
            '</td>' +
            '<td style="vertical-align: middle; color: #000000;">' + phone + pager + '</td>' +
        '</tr>';
    }

    var celRow = "";
    if (mobile) {
        var cleanMobile = mobile.replace(/\D/g, "");
        if (cleanMobile.length === 10 || cleanMobile.length === 11) {
            cleanMobile = "55" + cleanMobile;
        } else if (cleanMobile.indexOf("55") !== 0 && cleanMobile.length > 8) {
            cleanMobile = "55" + cleanMobile;
        }
        var waLink = "https://wa.me/" + cleanMobile;

        celRow = '<tr>' +
            '<td style="vertical-align: middle; padding-right: 8px; width: 20px;">' +
                '<a href="' + waLink + '" target="_blank" style="text-decoration: none; cursor: pointer;">' +
                    '<img src="' + icoCel + '" width="16" height="16" alt="WhatsApp / Celular" style="border: 0; display: block;">' +
                '</a>' +
            '</td>' +
            '<td style="vertical-align: middle; color: #000000;">' +
                '<a href="' + waLink + '" target="_blank" style="color: #000000; text-decoration: none; cursor: pointer;" title="Conversar no WhatsApp">' + mobile + '</a>' +
            '</td>' +
        '</tr>';
    }

    var disclaimerPt = "Esta mensagem pode conter informa&ccedil;&otilde;es confidenciais e/ou privilegiadas. Se voc&ecirc; n&atilde;o for o destinat&aacute;rio ou a pessoa autorizada a receber esta mensagem, n&atilde;o pode usar, copiar ou divulgar as informa&ccedil;&otilde;es nela contidas ou tomar qualquer a&ccedil;&atilde;o baseada nessas informa&ccedil;&otilde;es. Se voc&ecirc; recebeu esta mensagem por engano, por favor, avise imediatamente o remetente, respondendo o e-mail, e em seguida apague-o. Agradecemos sua coopera&ccedil;&atilde;o.";
    var disclaimerEn = "This message may contain confidential and/or privileged information. If you are not the addressee or authorized to receive this for the addressee, you must not use, copy, disclose or take any action based on this message or any information herein. If you have received this message in error, please advise the sender immediately by reply e-mail and delete this message. Thank you for your cooperation.";

    return '<div style="font-family: Arial, sans-serif; font-size: 10pt; color: #000000;">' +
        '<table cellpadding="0" cellspacing="0" border="0" style="font-size: 15px; border-collapse: collapse; font-family: Arial, sans-serif;">' +
            '<tbody>' +
                '<tr>' +
                    '<td valign="top" style="padding: 0px 15px 0px 0px; vertical-align: top;">' +
                        '<a href="' + linkRede + '" target="_blank" style="text-decoration: none; cursor: pointer;">' +
                            '<img src="' + logoUrl + '" width="150" height="150" alt="' + (isBemFamilia ? "Bem Familia" : "Lider Atacadista") + '" style="border: 0; display: block; width: 150px; height: 150px;">' +
                        '</a>' +
                    '</td>' +
                    '<td valign="top" style="padding: 0px; vertical-align: top; font-family: Arial, sans-serif; line-height: 1.3;">' +
                        '<div style="font-family: \'Arial Black\', Gadget, sans-serif; font-size: 14pt; color: #000000; font-weight: bold; margin-bottom: 2px;">' +
                            name +
                        '</div>' +
                        (cargoDepto ? '<div style="font-family: Arial, sans-serif; font-size: 11pt; font-style: italic; color: #222222; margin-bottom: 4px;">' + cargoDepto + '</div>' : '') +
                        '<div style="font-family: Arial, sans-serif; font-size: 10pt; color: #222222; margin-bottom: 8px;">' +
                            subLine +
                        '</div>' +
                        '<table cellpadding="0" cellspacing="0" border="0" style="font-size: 11pt; font-family: Arial, sans-serif; line-height: 1.5;">' +
                            '<tr>' +
                                '<td style="vertical-align: middle; padding-right: 8px; width: 20px;">' +
                                    '<img src="' + icoMail + '" width="16" height="16" alt="E-mail" style="border: 0; display: block;">' +
                                '</td>' +
                                '<td style="vertical-align: middle;">' +
                                    '<a href="mailto:' + email + '" style="color: #0066cc; text-decoration: underline; cursor: pointer;">' + email + '</a>' +
                                '</td>' +
                            '</tr>' +
                            foneRow +
                            celRow +
                        '</table>' +
                    '</td>' +
                '</tr>' +
            '</tbody>' +
        '</table>' +
        '<div style="margin-top: 15px; font-family: Verdana, sans-serif; font-size: 8pt; color: #666666; font-style: italic; line-height: 1.2;">' +
            disclaimerPt + '<br><br>' +
            '<span lang="EN-US">' + disclaimerEn + '</span>' +
        '</div>' +
    '</div>';
}

// Vincula os nomes das funcoes aos eventos registrados no manifesto
Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
