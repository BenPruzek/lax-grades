# coding: utf-8

import os
import sys

def create_stateSpace_defString(stateMatrix, directory, baseFileName):
    """Output the string 'stateMatrix'=[...] (modelica syntax) and the dimension of the matrix.

    Input is a csv-file
    """
    sOut = 'constant Real ' + stateMatrix +'[:,:] = [' + '\n'
    filename = os.path.join(directory, baseFileName+'_'+stateMatrix+'.csv')
    with open(filename, 'r') as fa:
        line = fa.readline()
        numCols = line.count(',') + 1
        numRows = 0
        while line:
            line = line.replace(' ', '').strip() + '%' + '\n'
            sOut += line
            line = fa.readline()
            numRows += 1
    cnt = sOut.count('%')        
    sOut = sOut.replace('%',';',cnt-1)
    sOut = sOut.replace('%','')
    sOut = sOut + '];' + '\n'
    return sOut, (numRows,numCols)


def get_numPorts_from_log(directory, baseFileName):
    """Obtain the number of ports from the log-file."""
    filename = os.path.join(directory, baseFileName + '.log')
    with open(filename, 'r') as fa:
        line = fa.readline()
        while line:
            if 'dim(y)' in line:
                tmp = line.split('=')
                secPart = tmp[1].lstrip()
                tmp = secPart.split(' ')
                break
            line = fa.readline()
    return int(tmp[0])


def get_refImp_from_log(directory, baseFileName):
    """Obtain the reference impedance from the log-file.
    
    The reference impedance have to be identical for all ports.
    """
    filename = os.path.join(directory, baseFileName + '.log')
    numPorts = get_numPorts_from_log(directory, baseFileName)
    refImpedances = []
    cnt = 0
    lineNumberStart = -1
    with open(filename, 'r') as fa:
        line = fa.readline()
        while line:
            cnt += 1
            if lineNumberStart > 0:
                tmp = line.split(':')
                secPart = tmp[1].lstrip()
                tmp = secPart.split(' ')
                refImpedances.append(float(tmp[0].rstrip()))                
                if cnt >= lineNumberStart + numPorts:
                    break
            if 'Reference impedances' in line:
                lineNumberStart = cnt
            line = fa.readline()
    return refImpedances


def get_portLabels(directory, baseFileName):
    filename = os.path.join(directory, baseFileName + '.log')
    numPorts = get_numPorts_from_log(directory, baseFileName)
    portNames = []
    cnt = 0
    lineNumberStart = -1
    with open(filename, 'r') as fa:
        line = fa.readline()
        while line:
            cnt += 1
            if lineNumberStart > 0:
                tmp = line.split('"')
                name = tmp[1].lstrip().rstrip()
                name = name.replace("+","_pos")
                name = name.replace("-","_neg")
                if name == str(cnt - lineNumberStart):
                    portNames.append('p_' + name)
                else:
                    portNames.append('p_{}'.format(cnt - lineNumberStart) +'_' + name)
                if cnt >= lineNumberStart + numPorts:
                    break                    
            if 'Port labels' in line:
                lineNumberStart = cnt
            line = fa.readline()
    return portNames


def get_vectorizedAddBlock_string():
    """Outputs the definition text defining VectorizedAdd block."""
    
    sOut = '''encapsulated block VectorizedAdd "Output the sum of the two inputs"
    import Modelica.Blocks.Interfaces.MI2MO;
    extends MI2MO;

    parameter Real k1[n]=fill(1,n) "Gain of input signal 1";
    parameter Real k2[n]=fill(1,n) "Gain of input signal 2";

    equation
        y = k1 .* u1 + k2 .* u2;
    annotation (
        Icon(coordinateSystem(
        preserveAspectRatio=true,
        extent={{-100,-100},{100,100}}), graphics={
        Line(points={{-100,60},{-74,24},{-44,24}}, color={0,0,127}),
        Line(points={{-100,-60},{-74,-28},{-42,-28}}, color={0,0,127}),
        Ellipse(lineColor={0,0,127}, extent={{-50,-50},{50,50}}),
        Line(points={{50,0},{100,0}}, color={0,0,127}),
        Text(extent={{-38,-34},{38,34}}, textString="+"),
        Text(extent={{-100,52},{5,92}}, textString="%k1"),
        Text(extent={{-100,-92},{5,-52}}, textString="%k2")}),
        Diagram(coordinateSystem(preserveAspectRatio=true, extent={{-100,-100},{
          100,100}}), graphics={Rectangle(
          extent={{-100,-100},{100,100}},
          lineColor={0,0,127},
          fillColor={255,255,255},
          fillPattern=FillPattern.Solid),Line(points={{50,0},{100,0}},
        color={0,0,255}),Line(points={{-100,60},{-74,24},{-44,24}}, color={
        0,0,127}),Line(points={{-100,-60},{-74,-28},{-42,-28}}, color={0,0,127}),
        Ellipse(extent={{-50,50},{50,-50}}, lineColor={0,0,127}),Line(
        points={{50,0},{100,0}}, color={0,0,127}),Text(
          extent={{-36,38},{40,-30}},
          textString="+"),Text(
          extent={{-100,52},{5,92}},
          textString="k1"),Text(
          extent={{-100,-52},{5,-92}},
          textString="k2")}));
end VectorizedAdd;\n\n'''
    return sOut


def get_additional_elements_string():
    """Outputs the definition text placing the parts needed to translate signals to electrical quantities."""
    
    sOut = '''Modelica.Electrical.Polyphase.Sources.SignalCurrent signalCurrent(m=numPins)
    annotation (Placement(transformation(
        extent={{-10,-10},{10,10}},
        rotation=0,
        origin={-36,22})));
Modelica.Electrical.Polyphase.Sensors.VoltageSensor voltageSensor(m=numPins)
    annotation (Placement(transformation(
        extent={{-10,10},{10,-10}},
        rotation=270,
        origin={-60,6})));
VectorizedAdd calc_a(
    n=numPins,
    k1=k1_a,
    k2=k2_a) 
    annotation (Placement(transformation(extent={{-10,-10},{10,10}})));
VectorizedAdd calc_i(
    n=numPins,
    k1=k1_i,
    k2=k2_i)
    annotation (Placement(transformation(extent={{50,-16},{70,4}})));
Modelica.Electrical.Polyphase.Basic.Star star(m=numPins)
    annotation (Placement(
        transformation(
        extent={{-10,-10},{10,10}},
        rotation=0,
        origin={16,-26})));
Modelica.Electrical.Analog.Basic.Ground gnd
    annotation (Placement(
        transformation(
        extent={{-10,-10},{10,10}},
        rotation=90,
        origin={40,-26})));\n\n'''
    return sOut


def get_connections_string():
    """Outputs the definition text connecting the parts which translate signals to electrical quantities."""
    
    sOut = '''connect(calc_a.y, stateSpace.u)
    annotation (Line(points={{11,0},{18,0}},color={0,0,127}));
connect(stateSpace.y, calc_i.u1)
    annotation (Line(points={{41,0},{48,0}}, color={0,0,127}));
connect(calc_i.y, signalCurrent.i)
    annotation (Line(points={{71,-6},{80,-6},{80,40},{-36,40},{-36,34}}, color={0,0,127}));
connect(star.pin_n, gnd.p)
    annotation (Line(points={{26,-26},{30,-26}}, color={0,0,255}));
connect(voltageSensor.v, calc_a.u1)
    annotation (Line(points={{-49,6},{-12,6}},   color={0,0,127}));
connect(voltageSensor.v, calc_i.u2)
    annotation (Line(points={{-49,6},{-42,6},{-42,-12},{48,-12}},color={0,0,127}));
connect(voltageSensor.plug_n, star.plug_p)
    annotation (Line(points={{-60,-4},{-60,-26},{6,-26}},color={0,0,255}));
connect(signalCurrent.plug_n, star.plug_p)
    annotation (Line(points={{-26,22},{-26,-26},{6,-26}}, color={0,0,255}));
connect(signalCurrent.plug_p, voltageSensor.plug_p)
    annotation (Line(points={{-46,22},{-60,22},{-60,16}}, color={0,0,255}));
connect(signalCurrent.i, calc_a.u2)
    annotation (Line(points={{-36,34},{-36,40},{-20,40},{-20,-6},{-12,-6}}, color={0,0,127}));
connect(voltageSensor.plug_p, p)
    annotation (Line(points={{-60,16},{-60,40}}, color={0,0,255}));\n\n'''
    return sOut


def replace_stateSpaceMatrices_in_moFile(moFileFullPath, directory, baseFileName):
    """Replace the definition of state space matrices in existing *.mo file.
    
    This functionality enables to replace only the underlying matrices after a recalculation.
    It is especially handy in cases where the user has performed changes in the pin layout and does not want
    to overwrite them.
    """
    sOut = ''
    outerRegion = True
    with open(moFileFullPath, 'r') as fa:
        line = fa.readline()
        while line:
            if '// 0001' in line:
                outerRegion = False
                sOut += line
                for letter in ('A','B','C','D'):
                    stringPart, dimension = create_stateSpace_defString(letter, directory, baseFileName)
                    sOut += stringPart
            elif '// 0002' in line:
                outerRegion = True
            if outerRegion == True:
                sOut += line
            if 'constant Integer numPins' in line:
                old_dim = int( (line.split('=')[1]).split(';')[0] )
            line = fa.readline()
    if dimension[0] == old_dim:
        with open(moFileFullPath, 'w') as fa:
            fa.write( sOut )
    else:
        print('WARNING: Number of ports are differnt. The given *.mo file was not modified.')
    return sOut


def get_moFile_string(modelName, directory, baseFileName):
    """Outputs the string in modelica syntax defining the model 'modelName'."""
    
    sOut = 'model ' + modelName + ' "Behaviour model from CST S-Parameters"' + '\n'
    
    sOut += get_vectorizedAddBlock_string()
    sOut += '// 0001 begin matrix def' + '\n'
    for letter in ('A','B','C','D'):
        stringPart, dimension = create_stateSpace_defString(letter, directory, baseFileName)
        sOut += stringPart
    sOut += '// 0002 end matrix def' + '\n'
    
    numPorts = dimension[0]
    refImpedances = get_refImp_from_log(directory, baseFileName)
    
    sOut += 'Modelica.Blocks.Continuous.StateSpace stateSpace( A=A, B=B, C=C, D=D )' + '\n' 
    sOut += 'annotation (Placement(transformation(extent={{20,-10},{40,10}})));' + '\n'
    sOut += '\n'
    sOut += 'constant Integer numPins=' + str(numPorts) + ';' + '\n'
    sOut += 'constant Real[numPins] k1_a = fill(1,numPins);' + '\n'
    sOut += 'constant Real[numPins] k2_a = {'
    for imp in refImpedances[:-1]: 
        sOut += str(imp) + ','
    sOut += str(refImpedances[-1]) + '};' + '\n'
    sOut += 'constant Real[numPins] k1_i = {'
    for imp in refImpedances[:-1]: 
        sOut += str(-1./imp) + ','
    sOut += str(-1./refImpedances[-1]) + '};' + '\n'
    sOut += 'constant Real[numPins] k2_i = {'    
    for imp in refImpedances[:-1]: 
        sOut += str(1./imp) + ','
    sOut += str(1./refImpedances[-1]) + '};' + '\n'    

    sOut += '\n'
      
    sOut += get_additional_elements_string()
    
    nPinsLeft = round(numPorts/2+0.1)
    nPinsRight = numPorts - nPinsLeft
    distBetweenPins = 100
    heightCanvas = distBetweenPins*(nPinsLeft)

    portNames = get_portLabels(directory, baseFileName) #still has whitespaces
    
    #create pins
    for i in range(nPinsLeft):
        ymin = int(heightCanvas/2) - 10 - int(distBetweenPins/2) - i*distBetweenPins
        ymax = ymin + 20
        sOut += 'Modelica.Electrical.Analog.Interfaces.Pin ' + str(portNames[i]) + '\n'
        sOut += ' annotation (Placement(iconTransformation(extent={{{{-110,{0}}},{{-90,{1}}}}})));'.format(ymin, ymax) + '\n'
    for i in range(nPinsRight):
        ymin = int(heightCanvas/2) - 10 - int(distBetweenPins/2) - i*distBetweenPins
        ymax = ymin + 20
        sOut += 'Modelica.Electrical.Analog.Interfaces.Pin ' + str(portNames[i+nPinsLeft]) + '\n'
        sOut += ' annotation (Placement(iconTransformation(extent={{{{90,{0}}},{{110,{1}}}}})));'.format(ymin, ymax) + '\n'
    
    sOut += 'protected' + '\n'
    sOut += 'Modelica.Electrical.Polyphase.Interfaces.PositivePlug p(m=numPins)' + '\n'
    sOut += '    annotation (Placement(transformation(extent={{-70,30},{-50,50}})));' + '\n'
    
    sOut += 'equation' + '\n'
    
    sOut += get_connections_string()
    
    #create connections to pins
    for i in range(numPorts):
        sOut += 'connect({0}, p.pin[{1}]);'.format(portNames[i],i+1) + '\n'
    
    sOut += 'annotation (Icon(coordinateSystem(extent={{{{-100,-{0}}},{{100,{1}}}}})));'.format(int(heightCanvas/2),int(heightCanvas/2)) + '\n'
    sOut += '\n'
    sOut += 'end '+ modelName + ';' + '\n'
    return sOut


if __name__=='__main__':
    outputFile = sys.argv[1]
    outputDir = sys.argv[2]
    inputDir = sys.argv[3]
    keepFlag = int(sys.argv[4]) # 0 means False(default)
    # outputFile = os.path.join(os.getcwd(), 'out.mo')
    # outputDir = os.getcwd()
    # inputDir = os.path.join(os.getcwd(), 'ABCD')
    # keepFlag = 1
    ABCDcsvfiles = [item for item in os.listdir(inputDir) if '.csv' in item]
    for file in ABCDcsvfiles:
        if '_A.csv' in file:
            baseFileName = file.replace( '_A.csv' , '')

    # # tests
    # numPorts = get_numPorts_from_log(inputDir, baseFileName)
    # refImps = get_refImp_from_log(inputDir, baseFileName)
    # portLabels = get_portLabels(inputDir, baseFileName)
    
    # write
    filename = os.path.join(outputDir, outputFile)
    if keepFlag == 0:    
        with open(filename, 'w') as fa:
            fa.write( get_moFile_string(baseFileName, inputDir, baseFileName) )
    else:
        if os.path.isfile( filename ):
            replace_stateSpaceMatrices_in_moFile(filename, inputDir, baseFileName)
        else:
            with open(filename, 'w') as fa:
                fa.write( get_moFile_string(baseFileName, inputDir, baseFileName) )
