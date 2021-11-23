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
# Performance enhancements:
# Currently, we find() each token on the string to be searched.
# In the future, use one of those libs that does a combined search on multiple strings:
# https://github.com/jakerachleff/commentzwalter
# https://github.com/abusix/ahocorapy
# https://github.com/WojciechMula/pyahocorasick/
# https://github.com/Guangyi-Z/py-aho-corasick
# we would then take the search resut of each keyword and build a function
# to combine the results into a final True/False query result.
#

from lark import Lark, Transformer
import sys

class SQuery:
    # default keywords
    or_keywords = ('"or"i', '"או"')
    and_keywords = ('"and"i', '', '"וגם"')
    not_keywords = ('"not"i', '"ללא"')

    def __init__(self):
        # set default keywords
        self.set_keywords(SQuery.or_keywords, SQuery.and_keywords, SQuery.not_keywords)
        # location of keyewords in string during query
        self.loc = [] # search keywords location in text
        self.len = [] # length of search keywords

    #
    # Call this if you want to localize your and/or keywords
    #
    def compile_parser(self):
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
        """.format(or_kw=' | '.join(self.or_keywords), and_kw=' | '.join(self.and_keywords), not_kw=' | '.join(self.not_keywords))
        #print(parser_ebnf)
        return Lark(parser_ebnf, parser="lalr", start='q_or', lexer='standard')

    # if you want to translate and/or to a different language
    def set_keywords(self, or_kw, and_kw, not_kw):
        (self.or_keywords, self.and_keywords, self.not_keywords) = (or_kw, and_kw, not_kw)
        # Must recompile parser after keywords are updated
        self.squery_parser = self.compile_parser()

    # if you want to translate and/or to a different language
    def get_keywords(self):
        return ((self.or_keywords, self.and_keywords, self.not_keywords))

    #
    # Compile a query, return a query function that accepts a string and returns True or False
    #
    def compile(self, query):
        self.query = query
        try:
            parsed = self.squery_parser.parse(query)
            #print(parsed.pretty())
            return SQuery_transformer(self).transform(parsed)
        except:
            return None

    def outer_s_string(self, s):
        st = str(s)
        index = len(self.len)
        self.len.append(len(st))
        self.loc.append(None)
        def s_string(text):
            i = text.find(st)
            self.loc[index] = i
            return i >= 0
        return s_string

    def outer_s_and(self, *f):
        def s_and(text):
            return all([func(text) for func in f])
        return s_and

    def outer_s_or(self, *f):
        def s_or(text):
            return any([func(text) for func in f])
        return s_or

    def outer_s_not(self, func):
        def s_not(text):
            return not func(text)
        return s_not

class SQuery_transformer(Transformer):
    def __init__(self, squery):
        super().__init__()
        self.squery = squery

    def q_or(self, items):
        #print("q_or:", items)
        return self.squery.outer_s_or(*items)
    def q_and(self, items):
        #print("q_and:", items)
        return self.squery.outer_s_and(*items)
    def q_not(self, items):
        #print("q_not:", items[0])
        return self.squery.outer_s_not(items[0])
    def escaped_string(self, items):
        #print("string:", items)
        return self.squery.outer_s_string(items[0][1:-1])
    def string(self, items):
        #print("string:", items)
        return self.squery.outer_s_string(items[0])

#
# Test code
#
if __name__ == "__main__":
    sq = SQuery()
    # enable this if you wantr to localize your commands.
    #sq.set_keywords(('"or"i', '"או"'), ('"and"i', '"וגם"'), ('"not"i', '"ללא"'))
    text = "abc def ghi"
    print("Text:", text)
    for s in ("abc and ghi", "not def"):
        query = sq.compile(s)
        if query is None:
            print("Invalid query!")
        print("Query: {} Result: {}".format(s, query(text)))
        print(sq.loc)
    sys.exit(0)

    while True:
        print("Enter text (type 'exit' to exit):")
        text = input()
        if text == "exit":
            sys.exit(0)
        print("Enter query:")
        s = input()
        query = sq.compile(s)
        if query is None:
            print("Invalid query!")
        else:
            #print(query)
            print(query(text))

