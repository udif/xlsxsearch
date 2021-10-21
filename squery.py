#
# A simple query language for searching text strings
# it accepts any words connected by and/or and their Hebrew counterparts
# it also honors ()'s.
#
from lark import Lark, Transformer
import sys

or_keywords = ('"or"i', '"או"')
and_keywords = ('"and"i', '"וגם"')
not_keywords = ('"not"i', '"ללא"')

or_kw_str = ""
and_kw_str = ""
not_kw_str = ""
squery_parser = None

#
# Call this if you want to localize your and/or keywords
#
def squery_compile_parser(or_kw_str=or_kw_str, and_kw_str=and_kw_str, not_kw_str=not_kw_str):
    global squery_parser
    parser_ebnf = r"""
        ?q_or : q_and (ws ( {or_kw} )  ws q_and)*
        ?q_and: q_val (ws ( {and_kw} ) ws q_val)*
        ?q_val: escaped_string
              | string
              | ( {not_kw} ) q_not 
              | "(" q_or ")"
        q_not: q_val
        string : STRING
        escaped_string : ESCAPED_STRING
        ws : WS
        STRING: /\w+/
        %import common.ESCAPED_STRING
        %import common.WS
        %ignore WS

    """.format(or_kw=or_kw_str, and_kw=and_kw_str, not_kw=not_kw_str)
    #print(parser_ebnf)
    squery_parser = Lark(parser_ebnf, start='q_or')

# if you want to translate and/or to a different language
def squery_set_keywords(or_kw, and_kw, not_kw):
    global or_kw_str, and_kw_str
    or_kw_str = ' | '.join(or_kw)
    and_kw_str = ' | '.join(and_kw)
    not_kw_str = ' | '.join(not_kw)
    squery_compile_parser(or_kw_str, and_kw_str, not_kw_str)

# set default keywords
squery_set_keywords(or_keywords, and_keywords, not_keywords)

#
# Compile a query, return a query function that accepts a string and returns True or False
#
def squery_compile(query):
    try:
        if squery_parser == None:
            squery_compile_parser()
        parsed = squery_parser.parse(query)
        #print(parsed.pretty())
        return SQuery_transformer().transform(parsed)
    except:
        return None

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

def outer_s_not(func):
    def s_not(text):
        return not func(text)
    return s_not

class SQuery_transformer(Transformer):
    def q_or(self, items):
        #print("q_or:", items, list(filter(None, items)))
        return outer_s_or(*list(filter(None, items)))
    def q_and(self, items):
        #print("q_and:", items, list(filter(None, items)))
        return outer_s_and(*list(filter(None, items)))
    def q_not(self, items):
        #print("q_not:", items[0])
        return outer_s_not(items[0])
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
        query = squery_compile(s)
        if query is None:
            print("Invalid query!")
        else:
            #print(query)
            print(query(text))

