/**
 * Suplemento de Assinatura Automática - Lider Atacadista & Bem Família
 * Plataforma: Microsoft 365 (Novo Outlook, OWA e Outlook Clássico)
 * Hospedagem: GitHub Pages (Lider-Atacadista/outlook-signature-addin)
 */

Office.onReady();

function onNewMessageComposeHandler(event) {
    try {
        var userProfile = Office.context.mailbox.userProfile;
        var email = (userProfile.emailAddress || "").toLowerCase().trim();
        var displayName = userProfile.displayName || "";
        
        // Determina a marca pelo domínio do UPN / E-mail
        var isBemFamilia = (email.indexOf("@bemfamilia.ind.br") !== -1);
        
        // Tenta buscar os dados adicionais de cargo, telefone e setor no users.json
        fetch("users.json?v=" + new Date().getTime())
            .then(function(response) {
                if (!response.ok) {
                    throw new Error("HTTP " + response.status);
                }
                return response.json();
            })
            .then(function(usersCatalog) {
                var userData = usersCatalog[email] || {};
                
                var signatureHtml = buildSignatureHtml({
                    isBemFamilia: isBemFamilia,
                    email: email,
                    name: userData.name || displayName || "Colaborador",
                    title: userData.title || "",
                    department: userData.department || "",
                    address: userData.address || (isBemFamilia ? "Rod. SC 445, Km 06" : "Rua José Alberto Nunes, s/n - Humaitá de Cima"),
                    city: userData.city || (isBemFamilia ? "Içara - SC" : "Tubarão - SC"),
                    phone: userData.phone || "",
                    pager: userData.pager ? (" | Ramal: " + userData.pager) : "",
                    mobile: userData.mobile || ""
                });

                applySignature(signatureHtml, event);
            })
            .catch(function(err) {
                console.warn("Aviso ao carregar users.json, utilizando dados do perfil:", err);
                
                // Fallback com dados do perfil
                var fallbackHtml = buildSignatureHtml({
                    isBemFamilia: isBemFamilia,
                    email: email,
                    name: displayName || "Colaborador",
                    title: "",
                    department: "",
                    address: isBemFamilia ? "Rod. SC 445, Km 06" : "Rua José Alberto Nunes, s/n - Humaitá de Cima",
                    city: isBemFamilia ? "Içara - SC" : "Tubarão - SC",
                    phone: "",
                    pager: "",
                    mobile: ""
                });

                applySignature(fallbackHtml, event);
            });
    } catch (ex) {
        console.error("Exceção no tratador de eventos:", ex);
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
            // Obrigatório sinalizar a conclusão do evento para o Outlook
            event.completed();
        }
    );
}

function buildSignatureHtml(d) {
    var companyName = d.isBemFamilia ? "Bem Família" : "Lider Atacadista";
    
    // Imagens hospedadas diretamente no repositório GitHub Pages
    var baseUrl = "https://lider-atacadista.github.io/outlook-signature-addin/images/";
    var logoUrl = d.isBemFamilia 
        ? baseUrl + "Logo-BemFamilia.png" 
        : baseUrl + "logo.png";
    var linkRede = d.isBemFamilia 
        ? "https://www.instagram.com/deliciasbemfamilia" 
        : "https://www.instagram.com/lider_atacadista";

    var icoMail = baseUrl + "ico-mail.png";
    var icoTel = baseUrl + "ico-tel.png";
    var icoCel = baseUrl + "ico-cel.png";

    var cargoDepto = "";
    if (d.title && d.department) {
        cargoDepto = d.title + " | " + d.department;
    } else if (d.title) {
        cargoDepto = d.title;
    } else if (d.department) {
        cargoDepto = d.department;
    }

    var subLine = "<b>" + companyName + "</b>";
    if (d.address) subLine += " - <i>" + d.address + "</i>";
    if (d.city) subLine += " - <b>" + d.city + "</b>";

    var foneRow = "";
    if (d.phone) {
        foneRow = '<tr>' +
            '<td style="vertical-align: middle; padding-right: 8px; width: 20px;">' +
                '<img src="' + icoTel + '" width="16" height="16" alt="Telefone" style="border: 0; display: block;">' +
            '</td>' +
            '<td style="vertical-align: middle; color: #000000;">' + d.phone + d.pager + '</td>' +
        '</tr>';
    }

    var celRow = "";
    if (d.mobile) {
        celRow = '<tr>' +
            '<td style="vertical-align: middle; padding-right: 8px; width: 20px;">' +
                '<img src="' + icoCel + '" width="16" height="16" alt="Celular" style="border: 0; display: block;">' +
            '</td>' +
            '<td style="vertical-align: middle; color: #000000;">' + d.mobile + '</td>' +
        '</tr>';
    }

    return '<div style="font-family: Arial, sans-serif; font-size: 10pt; color: #000000;">' +
        '<table cellpadding="0" cellspacing="0" border="0" style="font-size: 15px; border-collapse: collapse; font-family: Arial, sans-serif;">' +
            '<tbody>' +
                '<tr>' +
                    '<td valign="top" style="padding: 0px 15px 0px 0px; vertical-align: top;">' +
                        '<a href="' + linkRede + '" target="_blank" style="text-decoration: none;">' +
                            '<img src="' + logoUrl + '" width="150" height="150" alt="' + companyName + '" style="border: 0; display: block; width: 150px; height: 150px;">' +
                        '</a>' +
                    '</td>' +
                    '<td valign="top" style="padding: 0px; vertical-align: top; font-family: Arial, sans-serif; line-height: 1.3;">' +
                        '<div style="font-family: \'Arial Black\', Gadget, sans-serif; font-size: 14pt; color: #000000; font-weight: bold; margin-bottom: 2px;">' +
                            d.name +
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
                                    '<a href="mailto:' + d.email + '" style="color: #000000; text-decoration: none;">' + d.email + '</a>' +
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
            'Esta mensagem pode conter informações confidenciais e/ou privilegiadas. Se você não for o destinatário ou a pessoa autorizada a receber esta mensagem, não pode usar, copiar ou divulgar as informações nela contidas ou tomar qualquer ação baseada nessas informações. Se você recebeu esta mensagem por engano, por favor, avise imediatamente o remetente, respondendo o e-mail, e em seguida apague-o. Agradecemos sua cooperação.<br><br>' +
            '<span lang="EN-US">This message may contain confidential and/or privileged information. If you are not the addressee or authorized to receive this for the addressee, you must not use, copy, disclose or take any action based on this message or any information herein. If you have received this message in error, please advise the sender immediately by reply e-mail and delete this message. Thank you for your cooperation.</span>' +
        '</div>' +
    '</div>';
}

// Vincula o nome da função à ação registrada no manifesto
Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
