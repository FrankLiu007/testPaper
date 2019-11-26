require 'mathtype_to_mathml'
print MathTypeToMathML::Converter.new(ARGV[0]).convert
exit