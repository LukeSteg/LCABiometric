from pptx.util import Inches
import pieChartFactory
import barChartFactory
import textFactory
#import lineChartFactory

def generate_pie_chart(slide,shape,tokens,fileDict):
    pf = pieChartFactory.pieChartFactory(slide,shape)
    arg_dict = {}
    for token in tokens:
        tkn_type = token.split(':')[0]
        tkn_value = token.split(':')[1]
        arg_dict[tkn_type] = tkn_value
        
    if('X' in arg_dict):
        pf.setX(Inches(float(arg_dict['X'])))
    if('Y' in arg_dict):
        pf.setY(Inches(float(arg_dict['Y'])))
    if('CX' in arg_dict):
        pf.setCX(Inches(float(arg_dict['CX'])))
    if('CY' in arg_dict):
        pf.setCY(Inches(float(arg_dict['CY'])))
    if 'BOOK' in arg_dict:
    	pf.setBook(int(arg_dict['BOOK']))
    if('COLUMN' in arg_dict):
        pf.getDataFromColumn(int(arg_dict['COLUMN']),pf.getFileFromDict(fileDict))
    
    return pf.generateShape()

def generate_bar_chart(slide,shape,tokens,fileDict):
    bf = barChartFactory.barChartFactory(slide,shape)
    arg_dict = {}
    for token in tokens:
        tkn_type = token.split(':')[0]
        tkn_value = token.split(':')[1]

        arg_dict[tkn_type] = tkn_value

    if('X' in arg_dict):
        bf.setX(Inches(float(arg_dict['X'])))
    if('Y' in arg_dict):
        bf.setY(Inches(float(arg_dict['Y'])))
    if('CX' in arg_dict):
        bf.setCX(Inches(float(arg_dict['CX'])))
    if('CY' in arg_dict):
        bf.setCY(Inches(float(arg_dict['CY'])))
    if 'BOOK' in arg_dict:
    	bf.setBook(int(arg_dict['BOOK']))
    if('COLUMN' in arg_dict):
        bf.getDataFromColumn(int(arg_dict['COLUMN']),bf.getFileFromDict(fileDict))
    
#   if('CATEGORIES' in arg_dict):
#        bf.setCategories(arg_dict['CATEGORIES'])
#    if('SERIES NAME' in arg_dict and 'SERIES DATA' in arg_dict):
#        bf.addSeries(arg_dict['SERIES NAME'],arg_dict['SERIES DATA'])
#    if('HAS LEGEND' in arg_dict):
#        bf.setCategories(bool(arg_dict['HAS LEGEND']))
   
    return bf.generateShape()
    pass

def generate_line_chart(slide,shape,tokens,fileDict):
    lf = lineChartFactory.lineChartFactory(slide,shape)

    arg_dict = {}
    for token in tokens:
        tkn_type = token.split(':')[0]
        tkn_value = token.split(':')[1]

        arg_dict[tkn_type] = tkn_value
        
    if('X' in arg_dict):
        lf.setX(Inches(float(arg_dict['X'])))
    if('Y' in arg_dict):
        lf.setY(Inches(float(arg_dict['Y'])))
    if('CX' in arg_dict):
        lf.setCX(Inches(float(arg_dict['CX'])))
    if('CY' in arg_dict):
        lf.setCY(Inches(float(arg_dict['CY'])))
    if 'BOOK' in arg_dict:
    	lf.setBook(int(arg_dict['BOOK']))
    
    return lf.generateShape()

def generate_text(slide,shape,tokens,fileDict):
    tf = textFactory.textFactory(slide,shape)

    arg_dict = {}
    for token in tokens:
        tkn_type = token.split(':')[0]
        tkn_value = token.split(':')[1]

        arg_dict[tkn_type] = tkn_value
        
    if('X' in arg_dict):
        tf.setX(Inches(float(arg_dict['X'])))
    if('Y' in arg_dict):
        tf.setY(Inches(float(arg_dict['Y'])))
    if('CX' in arg_dict):
        tf.setCX(Inches(float(arg_dict['CX'])))
    if('CY' in arg_dict):
        tf.setCY(Inches(float(arg_dict['CY'])))
    if 'BOOK' in arg_dict:
    	tf.setBook(int(arg_dict['BOOK']))
    if('COLUMN' in arg_dict):
        tf.getDataFromColumn(int(arg_dict['COLUMN']),tf.getFileFromDict(fileDict))
    if('VARIABLE' in arg_dict):
        tf.computeOutputVar(arg_dict['VARIABLE'])
    
    return tf.generateShape()


def parse(slide,shape0,shape,fileDict):
    frame = shape.text_frame
    
    text = frame.text.strip();
    print "parsing %s" % text    
    if(containsQueryString(text)):
        text = getQueryString(text)
        print 'Query String ' + text
        tokens = text.split(',');
        print type(tokens[0]);
        tokens = map(str,tokens)
        print type(tokens[0]);
        tokens = map(str.strip,tokens);
        tokens = map(str.upper,tokens);
        
        fig_type = tokens[0].split(':')[1]

        switch = {'PIE CHART':generate_pie_chart , 'BAR CHART':generate_bar_chart , 'LINE CHART':generate_line_chart , 'TEXT':generate_text }
         
        new_shape = switch[fig_type](slide,shape0,tokens,fileDict)

def containsQueryString(text):
    result = False
    if(('#{' in text) and ('}' in text)):
        if(text.index('#{') < text.index('}')):
            result = True
    return result

def getQueryString(text):
    #modify to create multiple query strings?
    if(('#{' in text) and ('}' in text)):
        startIndex = text.index('#{')
        endIndex = text.index('}')
        return text[startIndex+2:endIndex]
    else:
        print 'WARNING query string not found when possibly expected'
        return ''

