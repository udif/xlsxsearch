#
# A simple query language for searching text strings
# it accepts any words connected by and/or and their Hebrew counterparts
# it also honors ()'s.
#
from lark import Lark, Transformer
import sys

or_keywords = ('"or"i', '"או"')
and_keywords = ('"and"i', '"וגם"')

or_kw_str = ""
and_kw_str = ""
squery_parser = None

#
# Call this if you want to localize your and/or keywords
#
def squery_compile_parser(or_kw_str=or_kw_str, and_kw_str=and_kw_str):
    global squery_parser
    parser_ebnf = r"""
        ?q_or : q_and (ws ( {or_kw} )  ws q_and)*
        ?q_and: q_val (ws ( {and_kw} ) ws q_val)*
        ?q_val: escaped_string
              | string
              | "(" q_or ")"
        string : STRING
        escaped_string : ESCAPED_STRING
        ws : WS
        STRING: /\w+/
        %import common.ESCAPED_STRING
        %import common.WS
        %ignore WS

    """.format(or_kw=or_kw_str, and_kw=and_kw_str)
    #print(parser_ebnf)
    squery_parser = Lark(parser_ebnf, start='q_or')

# if you want to translate and/or to a different language
def squery_set_keywords(or_kw, and_kw):
    global or_kw_str, and_kw_str
    or_kw_str = ' | '.join(or_kw)
    and_kw_str = ' | '.join(and_kw)
    squery_compile_parser(or_kw_str, and_kw_str)

# set default keywords
squery_set_keywords(or_keywords, and_keywords)

#
# Compile a query, return a query function that accepts a string and returns True or False
#
def squery_compile(query):
    if squery_parser == None:
        squery_compile_parser()
    return SQuery_transformer().transform(squery_parser.parse(query))

# In the future, use one of those:
# https://github.com/jakerachleff/commentzwalter
# https://github.com/abusix/ahocorapy
# https://github.com/WojciechMula/pyahocorasick/
# https://github.com/Guangyi-Z/py-aho-corasick

def outer_s_string(s):
    def s_string(text):
        return s in text
    return s_string

def outer_s_and(*f):
    def s_and(text):
        for func in f:
            if not func(text):
                return False
        return True
    return s_and

def outer_s_or(*f):
    def s_or(text):
        for func in f:
            if func(text):
                return True
        return False
    return s_or

class SQuery_transformer(Transformer):
    def q_or(self, items):
        #print("q_or:", items, list(filter(None, items)))
        return outer_s_or(*list(filter(None, items)))
    def q_and(self, items):
        #print("q_and:", items, list(filter(None, items)))
        return outer_s_and(*list(filter(None, items)))
    def escaped_string(self, items):
        #print("string:", items)
        return outer_s_string(items[0][1:-1])
    def string(self, items):
        #print("string:", items)
        return outer_s_string(items[0])
    def ws(self, items):
        #print("ws:", items)
        return None

#
# Test code
#
if __name__ == "__main__":
    while True:
        print("Enter text (type 'exit' to exit):")
        text = input()
        if text == "exit":
            sys.exit(0)
        print("Enter query:")
        s = input()
        #print(squery_parser.parse(s).pretty())
        query = squery_compile(s)
        print(query(text))

