import collections
import os
import pickle
import win32com.client as win32

mail_list = ['sarah.i.fajerberg@accenture.com', 'nikolaj.thomsen@accenture.com', 'avigiel.sara.cohen@accenture.com', 
				'laura.m.jensen@accenture.com', 'sofie.moth@accenture.com', 'sofie.e.nergaard@accenture.com', 
				'ekaterina.shchurova@accenture.com', 'd.k.singh@accenture.com', 'a.dadarkar@accenture.com', 
				'kristian.a.hede@accenture.com', 'linas.zicius@accenture.com', 'anne.gron@accenture.com', 
				'helga.bjarnadottir@accenture.com', 'theis.m.viborg@accenture.com', 'kathrine.m.stone@accenture.com', 
				'pawel.walas@accenture.com', 'lisa.holm.hansen@accenture.com']

try:
	# Unpickles the queue containing email addresses
	with open('queue_email_addresses.pkl', 'rb') as inp:
		email_addresses = pickle.load(inp)

except:
	
	print('Pickled queue does not exist, creating new queue.')
	email_addresses = collections.deque()
	
	for email_address in mail_list:
		email_addresses.append(email_address)

	
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
curr_email_address = email_addresses.popleft() 	# Saves the email address that was first in the queue
email_addresses.append(curr_email_address) 		# Puts the email address last in the queue
mail.To = curr_email_address
mail.Subject = 'Technology Breakfast'

mail.HTMLBody = (
'<!DOCTYPE html>'
'<html>'
'<body>'
'<p>Hello!</p>'
'<p>Technology Breakfast is happening again this Friday and this time the baton falls in your hands. <br> If you are not able to bring breakfast this time around, please find another volunteer on your own &#128522; (It doesn&rsquo;t need to be someone from the TPB team)</p>'
'<p><strong>Here&rsquo;s what needs to be done for the breakfast:</strong></p>'
'<ul style="margin-top: 0cm;">'
'<li>Place and pay for the order before Friday morning, in the app &ldquo;Lagkagehuset+&rdquo;, so you can skip the line by hitting the in-store &lsquo;Ring&amp;Collect&rsquo; bell at the counter (order example attached). The order can be fetched from 7:30am onwards.</li>'
'<li>Fetch the ordered bread from Lagkagehuset in Carlsbergbyen (next to Carlsberg station)</li>'
'<li>Clear the red/wooden table (on wheels) in the back of the Technology area.</li>'
'<li>Get IKEA-knives, cardboard plates, tissues and cutting boards from the wooden box labelled &lsquo;Technology People Board&rsquo; in the closet next to the refrigerator on 4F. (sometimes they are places elsewhere in the kitchen)</li>'
'<li>Take out needed amount of butter/cheese/jam from the white plastic bag in the refrigerator and place it on the table together with Lagkagehuset goodies. (If something is missing buy some from Netto)</li>'
'<li>(you might need to go down to the canteen to fetch a few knives and spoons for the butter and jam)</li>'
'<li>Have breakfast ready at <strong>8:30am</strong>.</li>'
'<li>Remove and pack down everything at <strong>10:30am</strong> and clear the table.</li>'
'</ul>'
'<p>Remember to expense on the wbs no. provided for Technology People Board engagements.</p>'
'<p>Any other questions, please let me know. <br> Thanks a lot for volunteering for Technology Breakfast (by the people for the people &#128522;)!</p>'
'<p>Kind regards, <br> Sarah, on behalf of Technology People Board</p>'
'</body>'
'</html>')

# Attaching the file called Lagkagehuset_order.png - make sure that it is placed in the right location
current_path = os.getcwd()
attachment = os.path.join(current_path, 'Lagkagehuset_order.png')
mail.Attachments.Add(attachment)

print('Do you want to send an email to ', curr_email_address, '?')
send_email = input('Enter "yes" to continue. Enter anything but "yes" to cancel.')

if send_email == 'yes' or send_email == 'Yes':
	print('Sending email.')
	mail.Send()
else:
	print('Cancel sending email.')

# Pickles the queue to be used next time this script runs
with open('queue_email_addresses.pkl', 'wb') as output:
	pickle.dump(email_addresses, output, pickle.HIGHEST_PROTOCOL)

