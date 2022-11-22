import pandas as pd
import win32com.client as win32

dados = pd.read_excel(r'email.xlsx')

codigo = dados['Código']
alvara = dados['Alvará']
contrato = dados['Contrato social']
demo = dados['Demonstrativo de resultado']
lao = dados['LAO']
qaf = dados['QAF']
termo = dados['Resp. Social']
iso9 = dados['ISO 9001']
iso14 = dados['ISO 14001']
iso45 = dados['ISO 45001']
email = dados['E-mail']

master = []
alv = []
con = []
dr = []
l = []
q = []
ter = []
i9 = []
i14 = []
i45 = []

for n in range(0,100):
    alvarat = alvara.isnull()
    contratot = contrato.isnull()
    demot = demo.isnull()
    laot = lao.isnull()
    qaft = qaf.isnull()
    termot = termo.isnull()
    iso9t = iso9.isnull()
    iso14t = iso14.isnull()
    iso45t = iso45.isnull()
    if alvarat[n] == 1 or laot[n] == 1 or qaft[n] == 1 or termot[n] == 1:
        master.append(1)
    else:
        master.append(0)
    if alvarat[n] == 1:
        alv.append('Alvará de funcionamento')
    else:
        alv.append(' ')
    if laot[n] == 1:
        l.append('Licença ambiental de operação')
    else:
        l.append(' ')
    if qaft[n] == 1:
        q.append('Questionário de avaliação de fornecedor(em anexo para ser preenchido e enviado de volta)')
    else:
        q.append(' ')
    if termot[n] == 1:
        ter.append('Termo de responsabilidade social e ambiental(em anexo para ser assinado e enviado de volta)')
    else:
        ter.append(' ')
    if contratot[n] == 1:
        con.append('Contrato social')
    else:
        con.append(' ')
    if demot[n] == 1:
        dr.append('Demonstrativo de resultado de 2022')
    else:
        dr.append(' ')
    if iso9t[n] == 1:
        i9.append('ISO 9001')
    else:
        i9.append(' ')
    if iso14t[n] == 1:
        i14.append('ISO 14001')
    else:
        i14.append(' ')
    if iso45t[n] == 1:
        i45.append('ISO 45001')
    else:
        i45.append(' ')
s = 40
print(f'Fornecedor: {codigo[s]}')
print(f'Master: {master[s]}')
print(f'Alvara: {alv[s]}')
print(f'LAO: {l[s]}')
print(f'QAF: {q[s]}')
print(f'Termo: {ter[s]}')
print(f'Contrato social: {con[s]}')
print(f'Demonstrativo de resultado: {dr[s]}')
print(f'ISO 9001: {i9[s]}')
print(f'ISO 14001: {i14[s]}')
print(f'ISO 45001: {i45[s]}')
print(f'email: {email[s]}')

for n in range(0,1):
    if master[n] == 1:
        outlook = win32.Dispatch("Outlook.Application")

        # Global variable
        unit_picker = ''


        class SaveMail:
            def __init__(self, mail_body, codigo):
                self.mail = outlook.CreateItem(0)
                self.body = mail_body
                self.codigo = codigo
                self.subject = f'Documentação - Docol - {codigo[n]}'

            def new_mail(self, email):
                self.email = email
                self.mail.to = email[n]
                # self.mail.HTMLBody = self.body
                # Open the window with email text
                self.mail.Display()
                index = self.mail.HTMLbody.find('>', self.mail.HTMLbody.find('<body'))
                self.mail.HTMLbody = self.mail.HTMLbody[:index + 1] + self.body + self.mail.HTMLbody[index + 1:]
                self.mail.Save()


        class MailBody:

            def __init__(self, alv, l, q, ter, con, dr, i9, i14, i45):
                self.alv = alv
                self.l = l
                self.q = q
                self.ter = ter
                self.con = con
                self.dr = dr
                self.i9 = i9
                self.i14 = i14
                self.i45 = i45
                self.text = ""

            def text_block(self):
                alv = self.alv
                l = self.l
                q = self.q
                ter = self.ter
                con = self.con
                dr = self.dr
                i9 = self.i9
                i14 = self.i14
                i45 = self.i45
                self.text = """
                    <!DOCTYPE html>
<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" lang="en">

<head>
	<title></title>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
	<meta name="viewport" content="width=device-width, initial-scale=1.0">
	<!--[if mso]><xml><o:OfficeDocumentSettings><o:PixelsPerInch>96</o:PixelsPerInch><o:AllowPNG/></o:OfficeDocumentSettings></xml><![endif]-->
	<style>
		* {
			box-sizing: border-box;
		}

		body {
			margin: 0;
			padding: 0;
		}

		a[x-apple-data-detectors] {
			color: inherit !important;
			text-decoration: inherit !important;
		}

		#MessageViewBody a {
			color: inherit;
			text-decoration: none;
		}

		p {
			line-height: inherit
		}

		.desktop_hide,
		.desktop_hide table {
			mso-hide: all;
			display: none;
			max-height: 0px;
			overflow: hidden;
		}

		@media (max-width:520px) {
			.desktop_hide table.icons-inner {
				display: inline-block !important;
			}

			.icons-inner {
				text-align: center;
			}

			.icons-inner td {
				margin: 0 auto;
			}

			.row-content {
				width: 100% !important;
			}

			.mobile_hide {
				display: none;
			}

			.stack .column {
				width: 100%;
				display: block;
			}

			.mobile_hide {
				min-height: 0;
				max-height: 0;
				max-width: 0;
				overflow: hidden;
				font-size: 0px;
			}

			.desktop_hide,
			.desktop_hide table {
				display: table !important;
				max-height: none !important;
			}
		}
	</style>
</head>

<body style="background-color: #FFFFFF; margin: 0; padding: 0; -webkit-text-size-adjust: none; text-size-adjust: none;">
	<table class="nl-container" width="100%" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; background-color: #FFFFFF;">
		<tbody>
			<tr>
				<td>
					<table class="row row-1" align="center" width="100%" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt;">
						<tbody>
							<tr>
								<td>
									<table class="row-content stack" align="center" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; color: #000000; width: 500px;" width="500">
										<tbody>
											<tr>
												<td class="column column-1" width="100%" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; font-weight: 400; text-align: left; vertical-align: top; padding-top: 5px; padding-bottom: 5px; border-top: 0px; border-right: 0px; border-bottom: 0px; border-left: 0px;">
													<table class="paragraph_block block-1" width="100%" border="0" cellpadding="10" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; word-break: break-word;">
														<tr>
															<td class="pad">
																<div style="color:#000000;direction:ltr;font-family:Arial, Helvetica Neue, Helvetica, sans-serif;font-size:13px;font-weight:400;letter-spacing:0px;line-height:120%;text-align:left;mso-line-height-alt:15.6px;">
																	<p style="margin: 0; margin-bottom: 16px;">Caro fornecedor,</p>
																	<p style="margin: 0; margin-bottom: 16px;">A Docol demanda alguns documentos de seus fornecedores e alguns dos que temos da sua empresa estão desatualizados.</p>
																	<p style="margin: 0; margin-bottom: 16px;">Alguns documentos são imprescindíveis e outros opcionais. É interessante que enviem os opcionais também, caso possuam.</p>
																	<p style="margin: 0; margin-bottom: 16px;">Abaixo segue uma lista dos documentos&nbsp;<strong>imprescindíveis</strong> que estão desatualizados, seguida da lista dos documentos opcionais.</p>
																	<p style="margin: 0;">Documentos <strong>imprescindíveis</strong>:</p>
																</div>
															</td>
														</tr>
													</table>
													<table class="list_block block-2" width="100%" border="0" cellpadding="10" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; word-break: break-word;">
														<tr>
															<td class="pad">
																<ul start="1" style="margin: 0; padding: 0; margin-left: 20px; list-style-type: revert; color: #000000; direction: ltr; font-family: Arial, Helvetica Neue, Helvetica, sans-serif; font-size: 13px; font-weight: 400; letter-spacing: 0px; line-height: 120%; text-align: left;">
																	<li style="margin-bottom: 0px;">{}</li>
																	<li style="margin-bottom: 0px;">{}</li>
																	<li style="margin-bottom: 0px;">{}</li>
																	<li>{}</li>
																</ul>
															</td>
														</tr>
													</table>
													<table class="paragraph_block block-3" width="100%" border="0" cellpadding="10" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; word-break: break-word;">
														<tr>
															<td class="pad">
																<div style="color:#000000;direction:ltr;font-family:Arial, Helvetica Neue, Helvetica, sans-serif;font-size:13px;font-weight:400;letter-spacing:0px;line-height:120%;text-align:left;mso-line-height-alt:15.6px;">
																	<p style="margin: 0;">Documentos opcionais:</p>
																</div>
															</td>
														</tr>
													</table>
													<table class="list_block block-4" width="100%" border="0" cellpadding="10" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; word-break: break-word;">
														<tr>
															<td class="pad">
																<ul start="1" style="margin: 0; padding: 0; margin-left: 20px; list-style-type: revert; color: #000000; direction: ltr; font-family: Arial, Helvetica Neue, Helvetica, sans-serif; font-size: 13px; font-weight: 400; letter-spacing: 0px; line-height: 120%; text-align: left;">
																	<li style="margin-bottom: 0px;">{}</li>
																	<li style="margin-bottom: 0px;">{}</li>
																	<li style="margin-bottom: 0px;">{}</li>
																	<li style="margin-bottom: 0px;">{}</li>
																	<li>{}</li>
																</ul>
															</td>
														</tr>
													</table>
													<table class="paragraph_block block-5" width="100%" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; word-break: break-word;">
														<tr>
															<td class="pad">
																<div style="color:#101112;direction:ltr;font-family:Arial, Helvetica Neue, Helvetica, sans-serif;font-size:13px;font-weight:400;letter-spacing:0px;line-height:120%;text-align:left;mso-line-height-alt:15.6px;">
																	<p style="margin: 0; margin-bottom: 16px;">É importante frisar o caráter urgente dessa demanda. Sendo assim, peço que esse e-mail seja respondido com a documentação solicitada o quanto antes.</p>
																	<p style="margin: 0; margin-bottom: 16px;">Agradeço a atenção e conto a colaboração da sua empresa.</p>
																	<p style="margin: 0;">Att.,</p>
																</div>
															</td>
														</tr>
													</table>
												</td>
											</tr>
										</tbody>
									</table>
								</td>
							</tr>
						</tbody>
					</table>
					<table class="row row-2" align="center" width="100%" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt;">
						<tbody>
							<tr>
								<td>
									<table class="row-content stack" align="center" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; color: #000000; width: 500px;" width="500">
										<tbody>
											<tr>
												<td class="column column-1" width="100%" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; font-weight: 400; text-align: left; vertical-align: top; padding-top: 5px; padding-bottom: 5px; border-top: 0px; border-right: 0px; border-bottom: 0px; border-left: 0px;">
													<table class="icons_block block-1" width="100%" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt;">
														<tr>
															<td class="pad" style="vertical-align: middle; color: #9d9d9d; font-family: inherit; font-size: 15px; padding-bottom: 5px; padding-top: 5px; text-align: center;">
																<table width="100%" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt;">
																	<tr>
																		<td class="alignment" style="vertical-align: middle; text-align: center;">
																			<!--[if vml]><table align="left" cellpadding="0" cellspacing="0" role="presentation" style="display:inline-block;padding-left:0px;padding-right:0px;mso-table-lspace: 0pt;mso-table-rspace: 0pt;"><![endif]-->
																			<!--[if !vml]><!-->
																			<table class="icons-inner" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; display: inline-block; margin-right: -4px; padding-left: 0px; padding-right: 0px;" cellpadding="0" cellspacing="0" role="presentation">
																				<!--<![endif]-->
																				<tr>
																					<td style="vertical-align: middle; text-align: center; padding-top: 5px; padding-bottom: 5px; padding-left: 5px; padding-right: 6px;"><a href="https://www.designedwithbee.com/" target="_blank" style="text-decoration: none;"><img class="icon" alt="Designed with BEE" src="https://d15k2d11r6t6rl.cloudfront.net/public/users/Integrators/BeeProAgency/53601_510656/Signature/bee.png" height="32" width="34" align="center" style="display: block; height: auto; margin: 0 auto; border: 0;"></a></td>
																					<td style="font-family: Arial, Helvetica Neue, Helvetica, sans-serif; font-size: 15px; color: #9d9d9d; vertical-align: middle; letter-spacing: undefined; text-align: center;"><a href="https://www.designedwithbee.com/" target="_blank" style="color: #9d9d9d; text-decoration: none;">Designed with BEE</a></td>
																				</tr>
																			</table>
																		</td>
																	</tr>
																</table>
															</td>
														</tr>
													</table>
												</td>
											</tr>
										</tbody>
									</table>
								</td>
							</tr>
						</tbody>
					</table>
				</td>
			</tr>
		</tbody>
	</table><!-- End -->
</body>

</html>
                """

                return self.text.format(alv[n], l[n], q[n], ter[n], con[n], dr[n], i9[n], i14[n], i45[n])
        new_mail.SaveMail.Send()