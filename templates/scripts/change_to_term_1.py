with open('for_term_1.vbs', 'r') as f:
	data = f.readlines()
	word_list = []
	for line in data:
		if 'Range' in line or 'Rows' in line:
			print(line)
		# words = line.split('(')
		# print(words)
		
	# 	for word in words:
	# 		word = word.strip()
	# 		if word == 'Range' or word == 'Rows':
	# 			word_list.append(word)
	print(len(word_list))

		# for i in range(len(line)):
		# 	pass
		# 	print(line[i] + line[i + 1])
		# break
			# try:
			# 	if line[i] + line[i+1] == 'R[':
			# 		print(line[i:])
			# except IndexError:
			# 	pass