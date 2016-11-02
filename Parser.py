from pptx.util import Inches
import pieChartFactory
import barChartFactory
import textFactory
#import lineChartFactory

def generate_pie_chart(slide,shape,tokens,fileRef):
    pf = pieChartFactory.pieChartFactory(slide,shape)
    arg_dict = {}
    for token in tokens:
        tkn_type = token.split(':')[0]
        tkn_value = token.split(':')[1]
        arg_dict[tkn_type] = tkn_value
        
    if('X' in arg_dict):
        pf.setX(Inches(int(arg_dict['X'])))
    if('Y' in arg_dict):
        pf.setY(Inches(int(arg_dict['Y'])))
    if('CX' in arg_dict):
        pf.setCX(Inches(int(arg_dict['CX'])))
    if('CY' in arg_dict):
        pf.setCY(Inches(int(arg_dict['CY'])))
    if('COLUMN' in arg_dict):
        pf.getDataFromColumn(int(arg_dict['COLUMN']),fileRef)
    
    return pf.generateShape()

def generate_bar_chart(slide,shape,tokens,fileRef):
    bf = barChartFactory.barChartFactory(slide,shape)
    arg_dict = {}
    for token in tokens:
        tkn_type = token.split(':')[0]
        tkn_value = token.split(':')[1]

        arg_dict[tkn_type] = tkn_value

    if('X' in arg_dict):
        bf.setX(Inches(int(arg_dict['X'])))
    if('Y' in arg_dict):
        bf.setY(Inches(int(arg_dict['Y'])))
    if('CX' in arg_dict):
        bf.setCX(Inches(int(arg_dict['CX'])))
    if('CY' in arg_dict):
        bf.setCY(Inches(int(arg_dict['CY'])))
    if('COLUMN' in arg_dict):
        bf.getDataFromColumn(int(arg_dict['COLUMN']),fileRef)
    
#   if('CATEGORIES' in arg_dict):
#        bf.setCategories(arg_dict['CATEGORIES'])
#    if('SERIES NAME' in arg_dict and 'SERIES DATA' in arg_dict):
#        bf.addSeries(arg_dict['SERIES NAME'],arg_dict['SERIES DATA'])
#    if('HAS LEGEND' in arg_dict):
#        bf.setCategories(bool(arg_dict['HAS LEGEND']))
   
    return bf.generateShape()
    pass

def generate_line_chart(slide,shape,tokens,fileRef):
    lf = lineChartFactory.lineChartFactory(slide,shape)

    arg_dict = {}
    for token in tokens:
        tkn_type = token.split(':')[0]
        tkn_value = token.split(':')[1]

        arg_dict[tkn_type] = tkn_value
        
    if('X' in arg_dict):
        lf.setX(Inches(int(arg_dict['X'])))
    if('Y' in arg_dict):
        lf.setY(Inches(int(arg_dict['Y'])))
    if('CX' in arg_dict):
        lf.setCX(Inches(int(arg_dict['CX'])))
    if('CY' in arg_dict):
        lf.setCY(Inches(int(arg_dict['CY'])))
    
    return lf.generateShape()

def generate_text(slide,shape,tokens,fileRef):
    tf = textFactory.textFactory(slide,shape)

    arg_dict = {}
    for token in tokens:
        tkn_type = token.split(':')[0]
        tkn_value = token.split(':')[1]

        arg_dict[tkn_type] = tkn_value
        
    if('X' in arg_dict):
        tf.setX(Inches(int(arg_dict['X'])))
    if('Y' in arg_dict):
        tf.setY(Inches(int(arg_dict['Y'])))
    if('CX' in arg_dict):
        tf.setCX(Inches(int(arg_dict['CX'])))
    if('CY' in arg_dict):
        tf.setCY(Inches(int(arg_dict['CY'])))
    if('TEXT' in arg_dict):
	print arg_dict['TEXT']
        tf.setText(arg_dict['TEXT'])
    
    return tf.generateShape()


def parse(slide,shape0,shape,fileRef):
    frame = shape.text_frame
    
    text = frame.text.strip();
    print "parsing %s" % text    
    print "text[0:1] == %s" % text[0:1]
    print "text[-1] == %s" % text[-1]
    if( text[0:2] == '#{' and text[-1]=='}' ):
        text = text[2:-1]
        tokens = text.split(',');
        print type(tokens[0]);
        tokens = map(str,tokens)
        print type(tokens[0]);
        tokens = map(str.strip,tokens);
        tokens = map(str.upper,tokens);
        
        fig_type = tokens[0].split(':')[1]

        switch = {'PIE CHART':generate_pie_chart , 'BAR CHART':generate_bar_chart , 'LINE CHART':generate_line_chart , 'TEXT':generate_text }
         
        new_shape = switch[fig_type](slide,shape0,tokens,fileRef)


