# squery - Copyright (c) 2021 Udi Finkelstein
#
# Redistribution and use in source and binary forms, with or without
# modification, are permitted provided that the following conditions are met:
#
# 1. Redistributions of source code must retain the above copyright notice, this
#    list of conditions and the following disclaimer.
#
# 2. Redistributions in binary form must reproduce the above copyright notice,
#    this list of conditions and the following disclaimer in the documentation
#    and/or other materials provided with the distribution.

# THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
# AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
# IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
# DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE
# FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
# DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
# SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
# CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
# OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
# OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

# A simple query language for searching text strings
# it accepts any words connected by and/or and their Hebrew counterparts
# it also honors ()'s.
#
# TODOs:
#
# 1. Performance enhancements:
# Currently, we find() each token on the string to be searched.
# In the future, use one of those libs that does a combined search on multiple strings:
# https://github.com/jakerachleff/commentzwalter
# https://github.com/abusix/ahocorapy
# https://github.com/WojciechMula/pyahocorasick/
# https://github.com/Guangyi-Z/py-aho-corasick
# we would then take the search resut of each keyword and build a function
# to combine the results into a final True/False query result.
#
# 2. Token location indication
# construct an optional search function that would also return the position of tokens
# that were found in the text so they can be visually marked for the user.
# As this would probably impact performance we would probably have this as an option
# using a new separate set of closures.

from lark import Lark, Transformer
import sys

or_keywords = ('"or"i', '"או"')
and_keywords = ('"and"i', '', '"וגם"')
not_keywords = ('"not"i', '"ללא"')

squery_parser = None

#
# Call this if you want to localize your and/or keywords
#
def squery_compile_parser(or_kw, and_kw, not_kw):
    global squery_parser
    parser_ebnf = r"""
        ?q_or : q_and (({or_kw})  q_and)*
        ?q_and: q_val (({and_kw}) q_val)*
        ?q_val: escaped_string
              | string
              | ( {not_kw} ) q_not 
              | "(" q_or ")"
        q_not: q_val
        string : STRING
        escaped_string : ESCAPED_STRING
        STRING: /\w+/
        %import common.ESCAPED_STRING
        %import common.WS
        %ignore WS
    """.format(or_kw=' | '.join(or_kw), and_kw=' | '.join(and_kw), not_kw=' | '.join(not_kw))
    #print(parser_ebnf)
    squery_parser = Lark(parser_ebnf, parser="lalr", start='q_or', lexer='standard')

# if you want to translate and/or to a different language
def squery_set_keywords(or_kw, and_kw, not_kw):
    global or_keywords, and_keywords, not_keywords
    (or_keywords, and_keywords, not_keywords) = (or_kw, and_kw, not_kw)
    squery_compile_parser(or_kw, and_kw, not_kw)

# if you want to translate and/or to a different language
def squery_get_keywords():
    global or_keywords, and_keywords, not_keywords
    return ((or_keywords, and_keywords, not_keywords))

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

def outer_s_string(s):
    def s_string(text):
        return s in text
    return s_string

def outer_s_and(*f):
    def s_and(text):
        return all(f)
    return s_and

def outer_s_or(*f):
    def s_or(text):
        return any(f)
    return s_or

def outer_s_not(func):
    def s_not(text):
        return not func(text)
    return s_not

class SQuery_transformer(Transformer):
    def q_or(self, items):
        #print("q_or:", items)
        return outer_s_or(*items)
    def q_and(self, items):
        #print("q_and:", items)
        return outer_s_and(*items)
    def q_not(self, items):
        #print("q_not:", items[0])
        return outer_s_not(items[0])
    def escaped_string(self, items):
        #print("string:", items)
        return outer_s_string(items[0][1:-1])
    def string(self, items):
        #print("string:", items)
        return outer_s_string(items[0])

#
# Test code
#
if __name__ == "__main__":
    # enable this if you wantr to localize your commands.
    #squery_set_keywords(('"or"i', '"או"'), ('"and"i', '"וגם"'), ('"not"i', '"ללא"'))
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

