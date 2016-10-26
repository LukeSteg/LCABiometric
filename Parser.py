from pptx.util import Inches
import pieChartFactory
import barChartFactory
import textFactory
import lineChartFactory

def generate_pie_chart(slide,shape,tokens):
	pf = pieChartFactory.pieChartFactory(slide,shape)
	for token in tokens
		tkn_type = token.split(':')[0]
		tkn_value = token.split(':')[1]

		arg_dict[tkn_type] = tkn_value
		
	if('X' in arg_dict):
		pf.setX(Inches(arg_dict['X']))
	if('Y' in arg_dict):
		pf.setY(Inches(arg_dict['Y']))
	if('CX' in arg_dict):
		pf.setCX(Inches(arg_dict['CX']))
	if('CY' in arg_dict):
		pf.setCY(Inches(arg_dict['CY']))
	if('CATEGORIES' in arg_dict):
		pf.setCategories(arg_dict['CATEGORIES'])
	if('SERIES NAME' in arg_dict and 'SERIES DATA' in arg_dict):
		pf.addSeries(arg_dict['SERIES NAME'],arg_dict['SERIES DATA'])
	if('HAS LEGEND' in arg_dict):
		pf.setCategories(bool(arg_dict['HAS LEGEND']))
	
	return pf.generateShape()

def generate_bar_chart(slide,shape,tokens):
	bf = barChartFactory.barChartFactory(slide,shape)
	for token in tokens
		tkn_type = token.split(':')[0]
		tkn_value = token.split(':')[1]

		arg_dict[tkn_type] = tkn_value

	if('X' in arg_dict):
		bf.setX(Inches(arg_dict['X']))
	if('Y' in arg_dict):
		bf.setY(Inches(arg_dict['Y']))
	if('CX' in arg_dict):
		bf.setCX(Inches(arg_dict['CX']))
	if('CY' in arg_dict):
		bf.setCY(Inches(arg_dict['CY']))
	if('CATEGORIES' in arg_dict):
		bf.setCategories(arg_dict['CATEGORIES'])
	if('SERIES NAME' in arg_dict and 'SERIES DATA' in arg_dict):
		bf.addSeries(arg_dict['SERIES NAME'],arg_dict['SERIES DATA'])
	if('HAS LEGEND' in arg_dict):
		bf.setCategories(bool(arg_dict['HAS LEGEND']))
	
	return bf.generateShape()

def generate_line_chart(slide,shape,tokens):
	lf = lineChartFactory.lineChartFactory(slide,shape)

	for token in tokens
		tkn_type = token.split(':')[0]
		tkn_value = token.split(':')[1]

		arg_dict[tkn_type] = tkn_value
		
	if('X' in arg_dict):
		lf.setX(Inches(arg_dict['X']))
	if('Y' in arg_dict):
		lf.setY(Inches(arg_dict['Y']))
	if('CX' in arg_dict):
		lf.setCX(Inches(arg_dict['CX']))
	if('CY' in arg_dict):
		lf.setCY(Inches(arg_dict['CY']))
	
	return lf.generateShape()

def generate_text(slide,shape,tokens):
	tf = textFactory.textFactory(slide,shape)

	for token in tokens
		tkn_type = token.split(':')[0]
		tkn_value = token.split(':')[1]

		arg_dict[tkn_type] = tkn_value
		
	if('X' in arg_dict):
		tf.setX(Inches(arg_dict['X']))
	if('Y' in arg_dict):
		tf.setY(Inches(arg_dict['Y']))
	if('CX' in arg_dict):
		tf.setCX(Inches(arg_dict['CX']))
	if('CY' in arg_dict):
		tf.setCY(Inches(arg_dict['CY']))
	if('TEXT' in arg_dict):
		tf.setText(arg_dict['TEXT'])
	
	return tf.generateShape()


def parse(slide,shape0,shape):
	frame = shape.text_frame
	print "parsing %s" % frame.text	
	
	text = frame.text.strip();
	if( text[0:1] == '#{' and text[-1]=='}' ):
		text = text[2:-1]
		tokens = text.split(',');
		tokens = map(str.strip,tokens);
		tokens = map(str.upper,tokens);
		
		fig_type = tokens[0].split(:)[1]

		switch = {'PIE CHART':generate_pie_chart , 'BAR CHART':generate_bar_chart , 'LINE CHART':generate_line_chart , 'TEXT':generate_text }
	
		new_shape = switch[fig_type](slide,shape0,tokens)


	
