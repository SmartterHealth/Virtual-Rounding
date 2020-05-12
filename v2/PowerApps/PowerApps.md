# Virtual Rounding Power Apps Deployment Instructions

Special thanks to KiZAN for their help with these instructions.

## SharePoint List Permissions

1. Update the Virtual Rounding list so that all Providers have &quot;Edit&quot; access to the list items.
2. If you desire for Patients to be able to conduct &quot;Invites&quot; from their devices, Room Accounts will need &quot;Edit&quot; access to their list items. Note that doing this could be a PHI exposure avenue if Patient devices are not locked down to Kiosk mode; as users could leverage the Web Browser to navigate to the SharePoint site and view other rooms&#39; data.

## Initial Provider App Upload

1. Navigate to the Power Apps portal signed in as a user with a Premium Power Automate license assigned to them. Click on &quot;Import&quot; and Upload the VirtualRounding\_xxxxx.zip file.
2. Set &quot;Import Setup&quot; for App as &quot;Create as new&quot;

[![](RackMultipart20200512-4-8g3p57_html_7302a85831e601e4.png)](https://github.com/justinkobel/Virtual-Rounding/blob/master/v2/Documentation/Images/ProviderAppImport-CreateAsNew.png)

1. Each of the Flows within the &quot;Related Resources&quot; should be configured with &quot;Create as new&quot;. Note, that if you don&#39;t see &quot;Set Patient Name&quot;, do not worry, this has been removed in v2.1 of the Power App.

[![](RackMultipart20200512-4-8g3p57_html_b2778d02b680759a.png)](https://github.com/justinkobel/Virtual-Rounding/blob/master/v2/Documentation/Images/ProviderAppImport-CreateFlowsAsNew.png)

1. Create a SharePoint Connection and Mail Connection (if one doesn&#39;t already exist) respectively.
2. Click &quot;Import&quot;
3. You should see a success indicator upon completion.

[![](RackMultipart20200512-4-8g3p57_html_f77fe3f487575dc2.png)](https://github.com/justinkobel/Virtual-Rounding/blob/master/v2/Documentation/Images/ProviderAppImport-CreateSuccess.png)

1. Open the uploaded App in &quot;Edit&quot; mode, and click on the &quot;Data&quot; icon on the left hand side of the interface.
2. Expand the previous &quot;Virtual Rounding&quot; data source and click &quot;Remove. Expect to see a lot of warnings pop up in your Health indicator.

[![](RackMultipart20200512-4-8g3p57_html_f582701e96e29e14.png)](https://github.com/justinkobel/Virtual-Rounding/blob/master/v2/Documentation/Images/ProviderAppImport-RemoveConnection.png)

1. In the &quot;Data sources&quot; pane, expand out &quot;Connectors&quot;, and click on your SharePoint connector.

[![](RackMultipart20200512-4-8g3p57_html_2c53c1dbfdb8c321.png)](https://github.com/justinkobel/Virtual-Rounding/blob/master/v2/Documentation/Images/ProviderAppImport-AddConnectionStep1.png)

1. Connect to your Virtual Rounding site (the v2 one) and select the Virtual Rounding list. If this list was imported as &quot;Virtual Rounding&quot;, you&#39;re good to go! If you chose an alternative name; you will need to replace references to &#39;Virtual Rounding&#39; in the Power App formulas with the name of your list.
2. Share the App &quot;File... Share within the Power Apps studio&quot; with any Providers who will need access to it (this should be via AAD group for larger deployments). You will probably want to un-select the &quot;Send email&quot; prompt until you are ready to share it with the users.

## Updating Power Automate Flows

1. Navigate to [Flow](https://flow.microsoft.com/) and sign into the system.
2. Click &quot;Reset Meeting Link&quot; and edit this flow.
3. On the first &quot;Get item&quot; step, update the step to refer to your Virtual Rounding site/list. Type-ahead on this will likely not work since the Flow was imported, click the &quot;Limit Columns by View&quot; to confirm that you got the URL and List Names correct

[![](RackMultipart20200512-4-8g3p57_html_6cdcf4c9ffc03990.png)](https://github.com/justinkobel/Virtual-Rounding/blob/master/v2/Documentation/Images/ResetMeetingLink-GetItem.png)

1. Populate the client Secret, Application Id, Directory Id, Hours off UTC and Set Time Zone steps with the proper information.
2. Expand the &quot;Update Item&quot; activity. Replace the Site Address and List Name with your site&#39;s information.
3. Ensure the columns below are populated with the correct information below. There may be &quot;left over&quot; column names like FamilyInvited or &quot;Room\_x0020\_UPN&quot;, which you can ignore (and will eventually decide to go away from the editor)

[![](RackMultipart20200512-4-8g3p57_html_1cbafc8bb18b6aa0.png)](https://github.com/justinkobel/Virtual-Rounding/blob/master/v2/Documentation/Images/ResetMeetingLink-UpdateItem.png)

1. Save the Flow, and run a test from the Power App to confirm that patient data, meeting invites, etc. are refreshed inside of the underlying SharePoint Virtual Rounding list.
2. Click the &quot;Share Meeting Link&quot; Flow and edit the flow.
3. On the first &quot;Get item&quot; step, update the step to refer to your Virtual Rounding site/list. (See step 4)
4. Confirm that the &quot;Send an email notification&quot; has the proper message content for your organization. (Some organizations prefer to replace this step with a &quot;Send As&quot; action within Exchange Online, or another mass-mailing platform. The default action will have a &quot;From&quot; address of Send Grid and &quot;Power Apps and Flow&quot; that will frequently wind up in users&#39; Spam folders.)
5. Update the &quot;Update Item&quot; action like in step 9 above. The formula for &quot;Shared With&quot; should be: add(body(&#39;Get\_item&#39;)?[&#39;SharedWith&#39;], 1)
6. Click &quot;Save&quot; and test this Flow from the Power App (remember, check your Spam folder)

If you start to get errors on Flows when testing them, remember sometimes it&#39;s easier to create a new one (and use the awesome new &quot;Copy to clipboard&quot; for each Flow action), than try to get a Flow with messed up connections patched up.

## Patient App Upload

1. Navigate to the Power Apps portal signed in as a user with a Premium Power Automate license assigned to them. Click on &quot;Import&quot; and Upload the PatientJoin\_xxxxx.zip file.
2. Set &quot;Import Setup&quot; for App as &quot;Create as new&quot;
3. For Share Meeting Link, select &quot;Create as New&quot;, otherwise we&#39;ll overwrite all of our hard work down on the previous steps to get the Meeting Link flow matched into your environment. In doing so, give it a name to denote it&#39;s temporary and able to be deleted in following steps.

[![](RackMultipart20200512-4-8g3p57_html_3e1c61e45f84a387.png)](https://github.com/justinkobel/Virtual-Rounding/blob/master/v2/Documentation/Images/PatientJoin-ImportMeetingLinkFlow.png)

1. Update the SharePoint and Mail connections just like the Provider app.

[![](RackMultipart20200512-4-8g3p57_html_6ec73d4032c3d98d.png)](https://github.com/justinkobel/Virtual-Rounding/blob/master/v2/Documentation/Images/PatientJoin-ImportApp.png)

1. Click &quot;Import&quot;
2. You should see a success indicator upon completion.

[![](RackMultipart20200512-4-8g3p57_html_f77fe3f487575dc2.png)](https://github.com/justinkobel/Virtual-Rounding/blob/master/v2/Documentation/Images/ProviderAppImport-CreateSuccess.png)

1. Open the uploaded App in &quot;Edit&quot; mode, and click on the &quot;Data&quot; icon on the left hand side of the interface.
2. Expand the previous &quot;Virtual Rounding&quot; data source and click &quot;Remove. Expect to see a lot of warnings pop up in your Health indicator.

[![](RackMultipart20200512-4-8g3p57_html_f582701e96e29e14.png)](https://github.com/justinkobel/Virtual-Rounding/blob/master/v2/Documentation/Images/ProviderAppImport-RemoveConnection.png)

1. In the &quot;Data sources&quot; pane, expand out &quot;Connectors&quot;, and click on your SharePoint connector.

[![](RackMultipart20200512-4-8g3p57_html_2c53c1dbfdb8c321.png)](https://github.com/justinkobel/Virtual-Rounding/blob/master/v2/Documentation/Images/ProviderAppImport-AddConnectionStep1.png)

1. Connect to your Virtual Rounding site (the v2 one) and select the Virtual Rounding list. If this list was imported as &quot;Virtual Rounding&quot;, you&#39;re good to go! If you chose an alternative name; you will need to replace references to &#39;Virtual Rounding&#39; in the Power App formulas with the name of your list.
2. [what about deleting the other flow]
3. Share the App with your VirtualRoundingRooms group you used to track all of your Rooms in AAD. You will want to un-select the &quot;Send email&quot; prompt, as users will not be receiving email on their devices.

## Branding Updates (Optional)

You will likely elect to replace logos within the Apps with your organization&#39;s logo, and also adjust the app&#39;s color scheme.

## Embedding the App in Teams for Your Providers and Rooms

If you wish to embed these Apps below directly in Teams, follow the below steps. Otherwise, the apps may be interacted with via &quot;traditional&quot; Power Apps desktop applications, web browser links, mobile apps, etc. for providers and rooms.

1. On the Providers and Patients Power App, click &quot;Edit Settings&quot;, and enable &quot;Preload app for enhanced performance&quot;

[![](RackMultipart20200512-4-8g3p57_html_426a4440f03757f9.png)](https://github.com/justinkobel/Virtual-Rounding/blob/master/v2/Documentation/Images/PowerApps-PreloadingEnable.png)

1. For the Providers and Patients App, click &quot;Add to Teams&quot;. This will generate a .zip file for each.

![](RackMultipart20200512-4-8g3p57_html_49ac0cb03196381.gif)

1. Navigate to the &quot;Teams Admin Center... Teams Apps... Manage apps&quot;
2. Click &quot;Upload new app&quot; and upload each .zip file you just generated. If you are unable to upload apps from here, check your &quot;Org-wide app settings&quot; and confirm that &quot;Custom Apps&quot; are enabled.
3. Navigate to &quot;Permission Policies&quot; for your App Permission Policies, and enable the &quot;Patient Join&quot; app to be published under &quot;Tenant apps&quot;

[![](RackMultipart20200512-4-8g3p57_html_96d4aa34f75b9aa1.png)](https://github.com/justinkobel/Virtual-Rounding/blob/master/v2/Documentation/Images/PermissionPolicy-PatientJoin.png)

1. Go to your Virtual Rounding Room &quot;Setup Policy&quot; that you configured during the setup policy, and add the &quot;Patient Join&quot; app as a Power App to the profile.
2. If you wish to publish the Provider Power App, follow similar steps as above, but do not target the &quot;Rooms&quot; policies, but instead &quot;Provider&quot; policies you will need to create to support this scenario.