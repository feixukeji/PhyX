import re
from sympy import latex, sympify
from typing import Callable, Literal, NoReturn, Pattern, Type, TypeVar

r"""
#### This module is designed for LaTeX code that only uses ASCII characters.

Utility to transform LaTeX code into an AST for calculation.

Remember to verify that the code is simple enough. Only use `\left` and `\right`.
"""


# TokenType = Literal["backslash", "text", "number", "symbol", "empty"]
# Token = tuple[TokenType, str]
# ASTNode = dict[str, str | int | float | bool |
#                list['ASTNode'] | Type['ASTNode']]  # Classical JSON :)

rules= [
    (re.compile("^\\\\"), "backslash"),
    (re.compile("^[a-zA-Z]+"), "text"),
    (re.compile("^[0-9]+(\\.[0-9]+)?"), "number"),
    (re.compile("^[\n\t ]+"), "empty")
]


def never():
    raise Exception("ERROR!")


def getTokens(string):
    r"""
Return the tokens of a LaTeX text.

Each token is a tuple of its type and value. There are 5 types:
|name|meaning|
|:-:|:-:|
|backslash|`\`|
|text|plain text|
|number|inteager or float|
|symbol|other non-blank character|
|empty|blank characters|

The type of value is always `str`.

Return type:
```python
list[tuple[Literal["backslash", "text", "number", "symbol", "empty"], str]]
```
"""
    global rules
    tokens: list= []

    # region Generate Tokens
    while len(string) > 0:
        for r, t in rules:
            tmp = r.search(string)
            if tmp == None:
                continue
            tokens.append((t, string[:tmp.span()[1]]))
            string = string[tmp.span()[1]:]
            break
        else:
            tokens.append(("symbol", string[:1]))
            string = string[1:]
    # endregion
    # region Check Tokens
    for t, v in tokens:
        if t == "symbol":
            if "+-*/^_=%|()[]{}.,'\"".find(v) == -1:
                raise Exception(f"Unsupported symbol: '{v}'")

        elif t == "backslash" or t == "text" or t == "number" or t == "empty":
            continue
        else:
            never()
    # endregion
    return tokens


PRECEDENCE = {
    "const": 10000,
    "variable": 10000,
    "number": 10000,
    "subexpr": 10000,
    "polynomial": 1,
    "monimimal": 2,
    "power": 3,
    "function": 10000
}


def parseTokens(tokens):
    r"""
Generate a very complex abstract syntax tree using these tokens and return its root node.

Following are the types of AST nodes:

|type|meaning|content|
|:-:|:-:|:-:|
|const|immutable variable|its name|
|variable|variavle to calculate the uncertainty; only appears after generation|its name|
|number|a directly given value; eg. 1.23|its value|
|subexpr|a subexpression, eg. ()[]{}|its value|
|polynomial|add its children directly|list of its child|
|monomial|multiply its children directly|list of its child|
|power|power B of A (A^B)|a, b|
|function|function|its name|

All of the AST nodes must have "children" attribute.

There's no need to distinguish the types "const" and "number" while processing the AST.
"""

    
    tmp= []

    # region Early Process
    i = iter(tokens)

    while True:
        try:
            t, v = next(i)
        except StopIteration:
            break
        if t == "empty":
            pass

        # region Macro Processing
        elif t == "backslash":
            try:
                t1, v1 = next(i)
            except StopIteration:
                raise Exception("No macro after '\\'")

            if t1 == "text":
                if v1 == "left":
                    try:
                        while True:
                            t2, v2 = next(i)
                            if t2 != "empty":
                                break
                    except StopIteration:
                        raise Exception("No symbol after '\\left'")

                    if t2 == "symbol":
                        if v2 == "(" or v2 == "[":
                            tmp.append(("lbarket", v2))
                        elif v2 == ".":
                            raise Exception("DO NOT use '\\left.'")
                        elif v2 == "|":
                            tmp.append(("function", "abs"))
                            tmp.append(("lbarket", "|"))
                        else:
                            raise Exception(
                                "No supported symbol after '\\left'")

                    elif t2 == "backslash":
                        try:
                            t3, v3 = next(i)
                        except StopIteration:
                            raise Exception("No symbol after '\\left\\'")
                        if t3 == "symbol" and v3 == "{":
                            tmp.append(("lbarket", v3))
                        else:
                            raise Exception(
                                "No supported symbol after '\\left\\'")
                    else:
                        raise Exception("No supported symbol after '\\left\\'")

                elif v1 == "right":
                    try:
                        while True:
                            t2, v2 = next(i)
                            if t2 != "empty":
                                break
                    except StopIteration:
                        raise Exception("No symbol after '\\right'")

                    if t2 == "symbol":
                        if v2 == ")" or v2 == "]":
                            tmp.append(("rbarket", v2))
                        elif v2 == ".":
                            raise Exception("DO NOT use '\\right.'")
                        elif v2 == "|":
                            tmp.append(("rbarket", "|"))
                        else:
                            raise Exception(
                                "No supported symbol after '\\right'")

                    elif t2 == "backslash":
                        try:
                            t3, v3 = next(i)
                        except StopIteration:
                            raise Exception("No symbol after '\\right\\'")
                        if t3 == "symbol" and v3 == "}":
                            tmp.append(("lbarket", v3))
                        else:
                            raise Exception(
                                "No supported symbol after '\\right\\'")
                    else:
                        raise Exception(
                            "No supported symbol after '\\right\\'")

                elif (v1 == "sin" or v1 == "cos" or v1 == "tan" or v1 == "cot" or v1 == "sec" or v1 == "csc"
                      or v1 == "ln" or v1 == "exp" or v1 == "log" or v1 == "sqrt" or v1 == "abs"):
                    tmp.append(("function", v1))

                elif (v1 == "alpha" or v1 == "beta" or v1 == "gamma" or v1 == "lambda" or v1 == "delta" or
                      v1 == "Delta" or v1 == "phi" or v1 == "varphi" or v1 == "theta" or v1 == "vartheta" or
                      v1 == "sigma" or v1 == "nu" or v1 == "mu" or v1 == "epsilon" or v1 == "varepsilon" or v1 == "pi"):
                    tmp.append(("var", "\\"+v1))

                elif v1 == "cdot" or v1 == "times":
                    tmp.append(('symbol', "*"))

                elif v1 == "partial" or v1 == "differential":
                    tmp.append(("symbol", v1))

                elif v1 == 'frac':
                    tmp.append(("fraction", ""))

                else:
                    raise Exception(f"Unsupported macro: '\\{v1}'")

            elif t1 == "empty":
                continue
            elif t1 == "number":
                raise Exception("No macro after '\\'")
            elif t1 == "backslash":
                continue
            elif t1 == "symbol":
                if v1 == ';':
                    continue
                elif v1 == '|':
                    raise Exception("Please use \\left\\| and \\right\\|")
                elif v1 == "{" or v1 == "}":
                    raise Exception("Please use \\left\\{ and \\right\\}")
                elif v1 == '%':
                    tmp.append(("symbol", v1))
                else:
                    raise Exception(f"Unknown symbol: \\{v1}")
            else:
                raise Exception(f"Unknown token type: {t}")
        # endregion
        # region Other Variables
        elif t == "text":
            tmp.append(("var", v))
            pass
        # endregion
        # region Simple Number
        elif t == "number":
            if v.find(".") != -1:
                tmp.append(("number", float(v)))
            else:
                tmp.append(("number", int(v)))
        # endregion
        # region Symbols
        elif t == "symbol":
            if v == "[" or v == "{":
                tmp.append(("lexprbarket", v))
            elif v == "]" or v == "}":
                tmp.append(("rexprbarket", v))
            elif v == '(' or v == ")":
                raise Exception("Please use \\left\\( and \\right\\)")
            elif v == '_':
                tmp.append(("subscript", "_"))
            elif v == '^':
                tmp.append(("superscript", "^"))
            elif v == '=':
                raise Exception("Not an expression!")
            elif v == "+" or v == "-" or v == "*" or v == "/":
                tmp.append(("symbol", v))
            else:
                tmp.append(("symbol", v))
        # endregion
    del i
    # endregion
    # region Barket Check
    barketStack = []
    for t, v in tmp:
        if t == "lbarket":
            barketStack.append(v)
        elif t == "lexprbarket":
            barketStack.append("e"+v)
        elif t == "rbarket":
            try:
                m = barketStack.pop()
                if m == "(" and v == ")" or m == "[" and v == "]" or m == "{" and v == "}" or m == "|" and v == "|":
                    continue
            except IndexError:
                pass
            raise Exception("Barket not match!")
        elif t == "rexprbarket":
            try:
                m = barketStack.pop()
                if m == "e[" and v == "]" or m == "e{" and v == "}":
                    continue
            except IndexError:
                pass
            raise Exception("Barket not match!")
    if len(barketStack) != 0:
        raise Exception("Barket not match!")
    del barketStack
    # endregion
    # region Add Ellipical Multiplication Sign
    i = 1
    l = len(tmp)
    while i < l:
        if((tmp[i][0] == "var" or tmp[i][0] == "lbarket" or tmp[i][0] == "function" or tmp[i][0] == "fraction")
                and (tmp[i-1][0] == "var" or tmp[i-1][0] == "number" or tmp[i-1][0] == "rexprbarket")):
            tmp.insert(i, ("symbol", "*"))
        i += 1
    del i, l
    # endregion
    # print(tmp)

    def findBarket(begin: int) -> int:
        nonlocal tmp
        t = tmp[begin]

        if t[0] == "lbarket":
            pairt = "rbarket"
        elif t[0] == "lexprbarket":
            pairt = "rexprbarket"
        else:
            raise Exception("Not a left barket")

        if t[1] == "(":
            pairv = ")"
        elif t[1] == "[":
            pairv = "]"
        elif t[1] == "{":
            pairv = "}"
        elif t[1] == "|":
            pairv = "|"
        else:
            never()

        l = len(tmp)
        stack = 0
        while begin < l:
            if tmp[begin] == t:
                stack += 1
            elif tmp[begin] == (pairt, pairv):
                stack -= 1
            if stack == 0:
                return begin
            begin += 1
        never()

    def find(begin: int, end: int, ele) -> list[int]:
        nonlocal tmp
        res: list[int] = []
        while begin < end:
            if tmp[begin] == ele:
                res.append(begin)
            if tmp[begin][0] == "lbarket" or tmp[begin][0] == "lexprbarket":
                begin = findBarket(begin)
            begin += 1
        return res

    def getSubExpr(begin: int, end: int, jump2OperatorPrecedence: int = 0):
        nonlocal tmp
        if end <= begin:
            raise Exception("Illegal expression")
        if tmp[begin] == ("lexprbarket", "{") and findBarket(begin) == end - 1:
            begin += 1
            end -= 1
            jump2OperatorPrecedence = 0
        # region +-
        if jump2OperatorPrecedence <= 1:
            l = find(begin, end, ("symbol", "+")) + \
                find(begin, end, ("symbol", "-"))
            l.sort()
            if len(l) > 0:
                l.append(end)
                children = []
                last = begin
                for i in l:
                    if last == i:
                        continue
                    children.append(getSubExpr(last, i, 2))
                    last = i
                return {
                    "type": "polynomial",
                    "children": children
                }
            del l
        # endregion
        # region */
        if jump2OperatorPrecedence <= 2:
            l = find(begin, end, ("symbol", "*")) + \
                find(begin, end, ("symbol", "/"))
            l.sort()
            if len(l) > 0 or tmp[begin] == ("symbol", "-"):
                l.append(end)
                factors: list= []

                # TODO: To be normalized in the method "clarify()"
                divide: list[bool] = []
                normalized: bool = True
                if tmp[begin] == ("symbol", "-"):
                    begin += 1
                    factors.append({
                        "type": "number",
                        "value": -1,
                        "children": []
                    })
                    divide.append(False)
                elif tmp[begin] == ("symbol", "+"):
                    begin += 1

                last = begin
                for index in l:
                    if(last != begin):
                        b = (tmp[last][1] == "/")
                        divide.append(b)
                        normalized &= not b
                        last += 1
                    else:
                        divide.append(False)
                    factors.append(getSubExpr(last, index, 3))
                    last = index
                return {
                    "type": "monomial",
                    "children": factors,
                    "divide": divide,
                    "normalized": normalized
                }

            elif tmp[begin] == ("symbol", "+"):
                begin += 1
            del l
        # endregion
        # region ^
        if jump2OperatorPrecedence <= 3:
            l = find(begin, end, ("superscript", "^"))
            if len(l) >= 1:
                return{
                    "type": "power",
                    "children": [
                        getSubExpr(begin, l[0], 4),
                        getSubExpr(l[0]+1, end, 4)
                    ]
                }
        # endregion
        # region fraction
        if tmp[begin][0] == "fraction":
            if tmp[begin + 1][0] == "lexprbarket":
                if tmp[begin + 1][1] == "[":
                    raise Exception("Barket not match!")
                l = findBarket(begin + 1)
                a = getSubExpr(begin + 2, l)
                l += 1
            else:
                a = getSubExpr(begin + 1, begin + 2)
                l = begin + 2

            if tmp[l][0] == "lexprbarket":
                if tmp[l][1] == "[":
                    raise Exception("Barket not match!")
                r = findBarket(l)
                b = getSubExpr(l + 1, r)
            else:
                b = getSubExpr(l, l + 1)
            return {  # TODO: This is only an expedient!!!
                "type": "fraction",
                "children": [a, b]
            }
        # endregion
        # region ()
        if tmp[begin][0] == "lbarket" and findBarket(begin) == end - 1:
            return {
                "type": "subexpr",
                "children": [getSubExpr(begin+1, end-1)]
            }
        # endregion
        # region var/const/number
        if tmp[begin][0] == "number":
            return {
                "type": "number",
                "value": tmp[begin][1],
                "children": []
            }
        elif tmp[begin][0] == "var":
            l = find(begin, end, ("subscript", "_"))
            if len(l) == 0:
                return {
                    "type": "variable",
                    "name": tmp[begin][1],
                    "children": []
                }
            elif len(l) == 1:
                if l[0] != begin + 1:
                    raise Exception("Strange usage of '_'")
                if tmp[begin + 2][0] == "lexprbarket":
                    if tmp[begin + 2][1] == "[":
                        raise Exception("Barket not match!")
                    r = findBarket(begin + 2)
                    res = ""
                    i = begin + 3
                    while i < r:
                        res += str(tmp[i][1])
                        i += 1
                    return {
                        "type": "variable",
                        "name": tmp[begin][1] + "_{" + res+"}",
                        "children": []
                    }
                else:
                    return {
                        "type": "variable",
                        "name": tmp[begin][1] + "_" + str(tmp[begin+2][1]),
                        "children": []
                    }
            else:
                raise Exception("Strange usage of '_'")

        # endregion
        # region function
        if tmp[begin][0] == "function":
            # TODO: Parameter for functions. eg. \sqrt[3]{3} or \log[10]{2}
            return {
                "type": "function",
                "name": tmp[begin][1],
                "children": [getSubExpr(begin + 1, end, 3)]
            }
        # endregion
        raise Exception("Illegal expression")

    return getSubExpr(0, len(tmp))


def clarifyAST(ast):
    # TODO: clarify
    return ast


def TeX2AST(tex: str):
    r"""
Transfor a TeX Code into AST.
"""
    return clarifyAST(parseTokens(getTokens(tex)))




def DFSAST(ast, function):
    r"""
Traverse the abstract syntax tree in a calculation order (depth first).

Maybe useless. And there's no promise to its speed and stability.
"""
    l= []
    for node in ast["children"]:
        l.append(DFSAST(node, function))
    return function(ast, l)


def replaceVariable(ast, replacement):
    r"""
Replace some variable or const node using provided number.

The paramater "replacement" is a dict contains names of the variables to be replaced and their value.
"""
    def replace(node, children):
        if node["type"] == "variable" or node["type"] == "const":
            if node["name"] in replacement.keys():
                return {
                    "type": "number",
                    "value": replacement[node["name"]],
                    "children": []
                }
            else:
                return node.copy()
        else:
            res = node.copy()
            res["children"] = children
            return res
    return DFSAST(ast, replace)


def AST2TeX(ast) -> str:
    r"""
Transform AST generated by this library to TeX Code.
"""
    def toString(node, strs: list[str]) -> str:
        if node["type"] == "const" or node["type"] == "variable":
            return node["name"]
        elif node["type"] == "number":
            return ('%.5g' % node["value"])
            # return latex(sympify(node["value"]).evalf(5), min=-3, max=3, mul_symbol="times")
        elif node["type"] == "subexpr":
            return f"\\left({strs[0]}\\right)"
        elif node["type"] == "polynomial":
            res: str = strs[0]
            i = 1
            while i < len(strs):
                if strs[i][0] == "-" or strs[i][0] == "+":
                    res += strs[i]
                else:
                    res += "+" + strs[i]
                i += 1
            return res
        elif node["type"] == "monomial":
            if strs[0] == "-1":
                res = "-" + strs[1]
                i = 2
            else:
                res: str = strs[0]
                i = 1
            while i < len(strs):
                if node["divide"][i]:
                    res += "/"+strs[i]
                else:
                    if ((node["children"][i]["type"] == "variable" or node["children"][i]["type"] == "const" or node["children"][i]["type"] == "subexpr" or
                         node["children"][i]["type"] == "function" or node["children"][i]["type"] == "fraction") and
                            (node["children"][i-1]["type"] == "variable" or node["children"][i-1]["type"] == "const" or node["children"][i-1]["type"] == "number")):
                        res += " "+strs[i]
                    else:
                        res += "\\times "+strs[i]
                i += 1
            return res
        elif node["type"] == "power":
            return f"{strs[0]}^{strs[1]}"
        elif node["type"] == "function":
            return f"\\{node['name']}{'{'}{strs[0]}{'}'}"
        elif node["type"] == "fraction":
            return "\\frac{" + strs[0] + "}{" + strs[1] + "}"

    return DFSAST(ast, toString)
