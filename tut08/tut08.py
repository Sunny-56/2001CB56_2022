
from datetime import datetime

start_time = datetime.now()
with open('scorecard.txt','a') as f:
	f.write('PAK ININNING			{}-{}\n\n'.format(147,10))



#Help
def scorecard(text=open('india_inns2.txt','r').read()):
	
	docts=[]
	for r in text.split('\n\n'):
		docts.append(r)
	#print(docts)


	import re
	

	ball=re.findall(r'\n([0-9].[0-6]|[1-2][0-9].[1-6])',text)
	#print(ball)
	# print(len(ball))
	# print(len(docts))
	batsman=[]
	baller=[]
	line=[]
	lst_out=[]
	for i in docts:
		line.append(i.split(","))
	# print(line[0][1])
	# print(line[0])
	# print(docts[1])
	for i in range(len(line)):
		#print([i][1])
		x=line[i][0].split(" ",1)
		
		
		player=x[1].split('to')
		
		batsman.append(player[1])
		baller.append(player[0])
	

	batsman=list(dict.fromkeys(batsman))
	baller=list(dict.fromkeys(baller))
	# for i in line:
	# 	print(i[1])
	lst_player=[]
	for i in line:
		x=i[0].split(" ",1)		
		player=x[1].split('to')
		#lst_player.append(player)
		temp_rn1=i[1].split('!!',1)
		# print(temp_rn1)
		temp_rn2=temp_rn1[0].split(' ',4)
		#print(temp_rn2)
		if((temp_rn2[1]=='out') and (temp_rn2[2]=='Bowled')):
			#print('c '+temp_rn2[-1]+'b '+player[0])
			lst_out.append('b '+' '+ player[0])
		if((temp_rn2[1]=='out') and (temp_rn2[2]=='Caught')):
			lst_out.append('c '+temp_rn2[-1]+' b '+player[0])
		if((temp_rn2[1]=='out') and (temp_rn2[2]=='Lbw')):
			
			lst_out.append('Lbw'+' '+'b '+' '+ player[0])
	print(lst_out)
	print(batsman)
	#CREATING LIST OF STASUS OF BATSMAN
	
	lst_six=[]
	lst_four=[]
	lst_run=[]
	lst_ball=[]
	lst_str=[]
	tot_run=0
	extra=0
	for i in range(len(batsman)):
		run=0
		six=0
		four=0
		ball=0
		for j in line:
			x=j[0].split(" ",1)		
			player=x[1].split('to')
			
			temp_rn1=j[1].split('!!',1)
			
			temp_rn2=temp_rn1[0].split(' ',4)
			if(batsman[i]==player[1]):
				if('1' in temp_rn2):
					run+=1
					ball+=1
				if(('2'in temp_rn2) and ('runs' in temp_rn2)):
					run+=2
					ball+=1
				if(('3' in temp_rn2) and ('runs' in temp_rn2)):
					run+=3
					ball+=1
				if('SIX' in temp_rn2):
					run+=6
					six+=1
					ball+=1
				if('FOUR' in temp_rn2):
					run+=4
					four+=1
					ball+=1
				if(('out' in temp_rn2) or ('no' in temp_rn2)or ('leg' in temp_rn2)):
					ball+=1
				if(('wide' in temp_rn2)):
					extra+=1
				if('wides' in temp_rn2):
					extra+=int(temp_rn2[-2])

		lst_four.append(four)
		lst_run.append(run)
		lst_six.append(six)
		lst_ball.append(ball)
		lst_str.append(round((run/ball)*100,2))


		
	for i in range(0,len(batsman)-len(lst_out)):
		lst_out.append('not out')
	
	

	
	print(baller)
	#CREATING LIST FOR ALL TH STATUS OF BALLER
	lst_med=[]
	lst_over=[]
	lst_wic=[]
	lst_no_ball=[]
	lst_wide=[]
	lst_eco=[]
	lst_Ball=[]
	lst_b_run=[]
	extra=0
	for i in range(len(baller)):
		over=0
		meden=0
		wic=0
		wide=0
		no_ball=0
		run=0
		ball=0
		
		for j in line:
			x=j[0].split(" ",1)		
			player=x[1].split('to')
			# print(player,'\n')
			temp_rn1=j[1].split('!!',1)
			# print(temp_rn2)
			temp_rn2=temp_rn1[0].split(' ',4)
			if(baller[i]==player[0]):
				if('1' in temp_rn2):
					run+=1
					ball+=1
				if(('2'in temp_rn2) and ('runs' in temp_rn2)):
					run+=2
					ball+=1
				if(('3' in temp_rn2) and ('runs' in temp_rn2)):
					run+=3
					ball+=1
				if('SIX' in temp_rn2):
					run+=6
					
					ball+=1
				if('FOUR' in temp_rn2):
					run+=4
					
					ball+=1
				if('out' in temp_rn2):
					ball+=1
					wic+=1
				if(('no' in temp_rn2)or ('leg' in temp_rn2)):
					ball+=1
				if(('wide' in temp_rn2)):
					run+=1
					wide+=1
					extra+=1
				if('wides' in temp_rn2):
					extra+=int(temp_rn2[-2])
					run+=int(temp_rn2[-2])
					wide+=int(temp_rn2[-2])
#APPENDING THE BALLER STATUS IN LIST

		over=round((ball/6),0)
		lst_over.append(over)
		lst_wic.append(wic)
		lst_eco.append(round((run/over),2))
		lst_wide.append(wide)
		lst_b_run.append(run)
		lst_med.append(meden)
	

	print(extra)
	total_run=sum(lst_run)+extra
	# print(total_run)
	new_lst_out=[]
	for i in lst_out:
		if(i!='not out'):
			new_lst_out.append(i)

#WRITE THE SCORE IN TEXT FILE
	with open('Scorecard.txt', 'a') as f:
		f.write(f"{'Batter' : <30}{'	' :^40}{'R': ^15}{'B': ^10}{'4s': ^10}{'6s': ^10}{'SR': >10}'\n")
		for i in range(len(batsman)):
			f.write(f"{batsman[i]: <30}{lst_out[i]: ^40}{lst_run[i]: ^15}{lst_ball[i]: ^10}{lst_six[i]: ^10}{lst_four[i]: ^10}{lst_str[i]: >10}'\n")
		f.write(f"\n{'Extras': <35}{extra+5: >20}({len(new_lst_out)}wkts)\n")
		# f.write(f"{'Total': <35}{total_run+5: >20}'\n")
		

		f.write(f"\n{'fall of wicket'}\n")
		f.write(f"\n{'Bowller' : <50}{'O' :^40}{'M': ^10}{'R': ^10}{'W': ^10}{'WD': ^10}{'ECO': >10}'\n")
		for i in range(len(baller)):
			f.write(f"{baller[i]: <50}{lst_over[i]: ^40}{lst_med[i]: ^10}{lst_b_run[i]: ^10}{lst_wic[i]: ^10}{lst_wide[i]: ^10}{lst_eco[i]: >10}'\n")
		f.write('\n\n\n\n\n')
		




scorecard(text=open('pak_inns1.txt','r').read())
with open('scorecard.txt','a') as f:
	f.write('INDIA ININNING					{}-{}\n\n'.format(148,5))

scorecard()

#This shall be the last lines of the code.
end_time = datetime.now()
print('Duration of Program Execution: {}'.format(end_time - start_time))
