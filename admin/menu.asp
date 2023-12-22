<!--#include file="../connections/bbg_conn.asp" -->
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>Quiz admin log-in</TITLE>

<link rel="stylesheet" href="styles/adminquizstyle.css" type="text/css">
</HEAD>
<BODY BGCOLOR=#FFCC00 leftmargin="0" topmargin="0" TEXT=#FFFFFF LINK=#FFFFFF VLINK=#FFFFFF alink="#FFFFFF">
<table width="150" border="0" cellspacing="0" cellpadding="2" background="images/admin01.gif" height="100%">
  <tr>
    <td class="menutext" height="20" colspan="2"><img src="images/quiz.gif" width="130" height="22"></td>
  </tr>
  <tr>
    <td class="menutext" height="3" colspan="2"><img src="images/hl.gif" width="130" height="3"></td>
  </tr>
  <tr>
    <td class="menutext" height="20" width="20"><img src="images/edit.gif" width="16" height="15"></td>
    <td class="menutext" height="20" width="130"><a href="q_list_of_subjects.asp" target="main">Edit pages </a></td>
  </tr>
  <tr>
    <td class="menutext" height="20"><img src="images/cert.gif" width="16" height="16"></td>
    <td class="menutext" height="20"><a href="q_certification_report.asp" target="main">Certification report</a></td>
  </tr>
  <!--
  <tr>
    <td class="menutext" height="20"><img src="images/new3.gif" width="16" height="15"></td>
    <td class="menutext" height="20"><a href="q_question_add.asp" target="main">Add
      question</a></td>
  </tr>-->
   <tr>
    <td class="menutext" height="20">&nbsp;</td>
    <td class="menutext" height="20">&nbsp;</td>
  </tr>
  <tr>
    <td class="menutext" height="20" colspan="2"><img src="images/rub_users.gif" alt=""></td>
  </tr>
  <tr>
    <td class="menutext" height="3" colspan="2"><img src="images/hl.gif" width="130" height="3"></td>
  </tr>
  <tr>
    <td class="menutext" height="20"><img src="images/nusers.gif" width="16" height="16"></td>
    <td class="menutext" height="20"><a href="q_list_of_users.asp" target="main">User
      results</a></td>
  </tr>
  <tr>
    <td class="menutext" height="20"><img src="images/cmpusers.gif" width="16" height="16"></td>
    <td class="menutext" height="20"><a href="q_comp_list_of_users.asp" target="main">Combined User Results</a></td>
  </tr>
      <tr>
    <td class="menutext" height="20"><img src="images/report.png" width="16" height="16"></td>
    <td class="menutext" height="20"><a href="q_list_of_subjects_stats.asp" target="main">Questions Stats</a></td>
  </tr>
  <tr>
    <td class="menutext" height="20"><img src="images/add.gif" width="16" height="16"></td>
    <td class="menutext" height="20"><a href="q_user_add.asp" target="main">Add New User</a></td>
  </tr>
  <tr>
    <td class="menutext" height="20"><img src="images/unread.gif" width="16" height="12"></td>
    <td class="menutext" height="20"><a href="email_report.asp" target="main">Email Report</a></td>
  </tr> 
   <!--tr>
    <td class="menutext" height="20"><img src="images/add.gif" width="16" height="16"></td>
    <td class="menutext" height="20"><a href="set_user_subjects.asp" target="main">Set user subjects</a></td>
  </tr
  <tr>
    <td class="menutext" height="20" colspan="2">&nbsp;</td>
  </tr>
  <tr>
    <td class="menutext" height="20" colspan="2"><img src="images/training.gif" width="130" height="17"></td>
  </tr>
  <tr>
    <td class="menutext" height="3" colspan="2"><img src="images/hl.gif" width="130" height="3"></td>
  </tr>
  <tr>
    <td class="menutext" height="20"><img src="images/edit.gif" width="16" height="15"></td>
    <td class="menutext" height="20"><a href="t_list_of_subjects.asp" target="main">Edit
      content</a></td>
  </tr>
  <tr>
    <td class="menutext" height="20"><img src="images/new3.gif" width="16" height="15"></td>
    <td class="menutext" height="20"><a href="t_question_add.asp" target="main">Add
      screens</a></td>
  </tr>-->
  <!-- <tr>
    <td class="menutext" height="20"><img src="images/monkey.gif" width="17" height="18"></td>
    <td class="menutext" height="20"><a href="t_list_of_monkeys.asp" target="main">Edit
      monkeys</a></td>
  </tr> -->
  <tr>
    <td class="menutext" height="20">&nbsp;</td>
    <td class="menutext" height="20">&nbsp;</td>
  </tr>
  <tr>
    <td class="menutext" height="20" colspan="2"><img src="images/bbg.gif" width="130" height="11"></td>
  </tr>
  <tr>
    <td class="menutext" height="3" colspan="2"><img src="images/hl.gif" width="130" height="3"></td>
  </tr>
  <tr>
    <td class="menutext" height="20"><img src="images/edit.gif" width="16" height="15"></td>
    <td class="menutext" height="20"><a href="b_list_of_subjects.asp" target="main">Edit
      content</a></td>
  </tr>
  <tr>
    <td class="menutext" height="20"><img src="images/new3.gif" width="16" height="15"></td>
    <td class="menutext" height="20"><a href="b_paragraph_add.asp" target="main">Add
      pages</a></td>
  </tr>
 <!-- <tr>
    <td class="menutext" height="20"><img src="images/hlp.gif" width="18" height="15"></td>
    <td class="menutext" height="20"><a href="b_list_of_help.asp" target="main">Edit
      HelpTabs</a></td>
  </tr>
  <tr>
    <td class="menutext" height="20"><img src="images/faq.gif" width="18" height="15"></td>
    <td class="menutext" height="20"><a href="b_list_of_faq.asp" target="main">Edit
      FAQTabs</a></td>
  </tr>-->
  <tr>
    <td class="menutext" height="20"><img src="images/reorder.gif" width="16" height="16"></td>
    <td class="menutext" height="20"><a href="b_list_of_replace.asp" target="main">Replacements</a></td>
  </tr>
   <tr>
      <td class="menutext" height="20">&nbsp;</td>
      <td class="menutext" height="20">&nbsp;</td>
  </tr>
	<tr>
	  <td class="menutext" height="20" colspan="2"><img src="images/ffr.gif" width="130" height="11"></td>
	</tr>
	<tr>
	  <td class="menutext" height="3" colspan="2"><img src="images/hl.gif" width="130" height="3"></td>
	</tr>
	<!--<tr>
	  <td class="menutext" height="20"><img src="images/edit.gif" width="16" height="15"></td>
	  <td class="menutext" height="20"><a href="f_list_activities.asp" target="main">Edit Field Feedback Options</a></td>
	</tr>-->
	<tr>
	  <td class="menutext" height="20"><img src="images/logs.gif" width="16" height="15"></td>
	  <td class="menutext" height="20"><a href="f_feedback_logs.asp" target="main">Field Feedback Reports</a></td>
  </tr>
  <tr>
    <td class="menutext" height="20">&nbsp;</td>
    <td class="menutext" height="20">&nbsp;</td>
  </tr>
	<tr>
	  <td class="menutext" height="3" colspan="2"><img src="images/settings.gif" alt=""></td>
	</tr>
	<tr>
	  <td class="menutext" height="3" colspan="2"><img src="images/hl.gif" width="130" height="3"></td>
	</tr>
    <tr>
    <td class="menutext" height="20"><img src="images/edit_email.gif" width="16" height="8"></td>
    <td class="menutext" height="20"><a href="edit_emails.asp" target="main">Edit Emails</a></td>
  </tr>
   <!--<tr>
    <td class="menutext" height="20"><img src="images/sending.gif" width="16" height="8"></td>
    <td class="menutext" height="20"><a href="send_business_division_email.asp" target="main">Send Business Division Email</a></td>
  </tr>
 <tr>
    <td class="menutext" height="20"><img src="images/auto_email.gif" width="16" height="12"></td>
    <td class="menutext" height="20"><a href="edit_auto_email.asp" target="main">Edit Auto Email</a></td>
  </tr>
  <tr>
    <td class="menutext" height="20"><img src="images/remind.gif" width="16" height="15"></td>
    <td class="menutext" height="20"><a href="edit_reminder_email.asp" target="main">Edit Manual Reminder Email</a></td>
  </tr>
   <tr>
    <td class="menutext" height="20"><img src="images/escalation.gif" width="16" height="15"></td>
    <td class="menutext" height="20"><a href="edit_escalation_email.asp" target="main">Edit Escalation<br>Email</a></td>
  </tr>
   <tr>
    <td class="menutext" height="20"><img src="images/edit_email.gif" width="16" height="13"></td>
    <td class="menutext" height="20"><a href="edit_final_email.asp" target="main">Edit Final Email</a></td>
  </tr>-->
  <tr>
    <td class="menutext" height="20"><img src="images/bus.gif" width="16" height="16"></td>
    <td class="menutext" height="20"><a href="business_level1.asp" target="main">Business</a></td>
  </tr>
  <tr>
    <td class="menutext" height="20"><img src="images/users.gif" width="16" height="16"></td>
    <td class="menutext" height="20">
      <p><a href="business_dpt.asp" target="main"><% =BBPinfo3 %></a></p>
    </td>
  </tr>
  <tr>
    <td class="menutext" height="20"><img src="images/setup.gif" width="16" height="16"></td>
    <td class="menutext" height="20">
      <p><a href="business_cmp.asp" target="main">Company</a></p>
    </td>
  </tr>
  <tr>
    <td class="menutext" height="20"><img src="images/pref.gif" width="16" height="16"></td>
    <td class="menutext" height="20"><a href="preferences.asp" target="main">Preferences</a></td>
  </tr>
  <tr>
    <td class="menutext" height="20"><img src="images/logs.gif" width="14" height="15"></td>
    <td class="menutext" height="20"><a href="logs.asp" target="main">Log files</a></td>
  </tr>
  <tr>
    <td class="menutext" height="20"><img src="images/logout.gif" width="16" height="16"></td>
    <td class="menutext" height="20"><a href="index.asp" target="_top">Logout</a></td>
  </tr>
  <tr>
    <td class="menutext" height="99%" colspan="2">&nbsp;</td>
  </tr>
</table>
</BODY>
</HTML>
