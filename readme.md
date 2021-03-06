<table border="0" cellpadding="0" cellspacing="0" width="95%">

<tbody>

<tr>

<td valign="top">

# <font face="VERDANA, ARIAL, HELVETICA" size="2">Read Me File for VC Turbo Sample Site</font>

**Commerce Server 3.0 sample site: Volcano Coffee Turbo**  
This document explains the operation of the VC Turbo sample site and provides technical details describing the implementation of the site using the Commerce Server 3.0 platform. We strongly encourage you to read the entire document before proceeding with installation.

**Contents**

- [Overview of VC Turbo](#overview)
- [Known issues](#issues)
- [Setup](#setup)
  - [Requirements](#requirements)
  - [Installing the VC Turbo site](#installing)
  - [Removing (uninstalling) the VC Turbo site](#removing)
- [Explore the site](#explore)

<font face="VERDANA, ARIAL, HELVETICA" size="2"><a name="top"></a><a name="overview"></a></font>

<font face="VERDANA, ARIAL, HELVETICA" size="2">

**Overview of VC Turbo**  
VC Turbo is a new version of the Volcano Coffee sample site that shipped with Site Server, Commerce Edition 3.0\. It has been modified to show how to improve performance and capacity for a Commerce Server site. VC Turbo forgoes certain features to achieve these improvements. Please see The Microsoft Site Server 3.0 Commerce Edition Performance Kit for the detailed documentation on these changes. (http://www.microsoft.com/siteserver/commerce/.) Volcano Coffee is documented in the Site Server 3.0 Commerce Edition documentation.

VC Turbo is a new addition to the set of Commerce Server 3.0 sample sites (the Clocktower, Volcano Coffee, Microsoft Press, Microsoft Market, and Trey Research sample sites were shipped with Site Server 3.0, Commerce Edition, and all except Trey Research are included in the Typical installation). It may be useful to compare the VC Turbo site to other sample sites to learn about differences and similarities.

The VC Turbo site can be copied using the Commerce Site Builder Wizard. For information, see Site Creation, Operation, and Administration in the Commerce Server 3.0 documentation.

_Note:_ VC Turbo, like other Commerce Server 3.0 sample sites, is intended to demonstrate features of Commerce Server 3.0. It can be used as the basis for creating a live production site.<a name="issues"></a>

**Known issues**

_Incompatibility between Microsoft FrontPage and Commerce pages_  
Microsoft FrontPage98 is not recommended for editing ASP files unless the files were originally created in FrontPage. Since the Commerce Server sample sites were not generated by FrontPage, editing them in FrontPage can result in loss of data and site functionality. For more information, see [http://www.microsoft.com/support/](http://www.microsoft.com/support/) You can, however, use Microsoft Visual InterDev 1.0 (a component of Microsoft Visual Studio 97). Visual InterDev 1.0 was used to develop the VC Turbo site. For information about editing a Commerce Server site in Visual InterDev, see How Do I... in the Commerce documentation.

_Unable to delete a site's directory_  
Occasionally, when removing VC Turbo using WebAdmin or the Commerce Host Administration snap-in for the Microsoft Management Console (MMC), an error message is displayed indicating that "Deleting files partially or completely failed." Clicking the Details button displays a message indicating that the directory is not empty. This error occurs when the snap-in fails to remove the VCTurbo directory and/or its content.

If this occurs, you should manually delete the VCTurbo directory. If you receive an "Error Deleting File" message when attempting to remove the VCTurbo directory, reboot your computer and then delete the directory.

_SSL Security_ To use SSL, the computer must have a server certificate obtained from a valid Certificate Authority (CA). By default, Secure HTTP is not enabled in the site so that you can use the site even if the computer does not have a certificate. To enable Secure HTTP, start the Commerce Host Administration snap-in for the Microsoft Management Console (MMC). In the scope pane, right-click the name of the site, and then click Properties on the shortcut menu. Click the Security tab, and then select the Enable HTTPS check box. For Secure host name, type the name of the server that you want to use for processing secure HTTP posts.<a name="setup"></a>

**Setup**<a name="requirements"></a>

_Requirements_

- To use the VC Turbo site, you must have already installed Site Server 3.0, Commerce Edition.
- The site is designed for use only with Microsoft® Internet Explorer 4.01 and Microsoft® SQL Server 6.5\.

<a name="installing"></a>

_Installing the VC Turbo site_  
The VC Turbo site is packaged as a self-extracting executable file. To install the site, double-click the VCTurbo.exe file to launch Setup, which installs the site into the appropriate directories. VC Turbo is installed in \inetpub\wwwroot\ or whatever directory you specified as the Web root during the installation of Microsoft Internet Information Server (IIS).

You must have already created a database for the sample site's files. During setup, you must select the name of the DSN to use for connecting to the database, and you must supply the database login name and password. It is recommended that you install VC Turbo in the same database as the one used for the rest of the sample sites during the Commerce Server 3.0 installation.<a name="removing"></a>

_Removing (uninstalling) the VC Turbo site_  
Use the Commerce Server administration tools (MMC, WebAdmin, or command-line) to remove the VC Turbo site. On the Start menu, point to Programs, point to Microsoft Site Server, point to Administration, and then click Site Server Service Admin (MMC). Expand Commerce Host Administration, select the host (server) on which the site is installed, and then click OK. In the scope (left) pane, right-click the VCTurbo folder, and then click Delete on the shortcut menu.<a name="explore"></a>

**Explore the site**  
Like the other Commerce Server sample sites, VC Turbo has both a catalog (shopping) site and Manager pages.

Shopping site: http://localhost/VCTurbo/

Manager pages: http://localhost/VCTurbo/manager/

When testing VC Turbo, you can use the following test credit-card numbers:

- American Express: 3111-1111-1111-1117
- Visa: 4111-1111-1111-1111
- MasterCard: 5111-1111-1111-1118
- Discover: 6111-1111-1111-1116

</font></td>

</tr>

</tbody>

</table>
