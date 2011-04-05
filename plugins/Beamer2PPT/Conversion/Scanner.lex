%x str, overlay, optional, pre_overlay, pre_optional

%using SimpleConverter.Contract;

%namespace SimpleConverter.Plugin.Beamer2PPT

ws          [ \t]*
wsp         [ \t]+
wsl         {ws}\r?\n?{ws}
envBegin    \\begin{wsl}
envEnd      \\end{wsl}

%{
    // global variables
    bool tabular = false;   // info if we are in tabular env., inclusive states don't seem to work and exclusive are out of question
%}

%%
%{
    // local variables
    string unformattedText = "";
    int spaces = 0;
    int nls = 0;
%}

// Parsing of LaTeX macros
// =============================================================================

// Preambule
// -----------------------------------------------------------------------------
\\documentclass { BEGIN(pre_optional); return (int) Tokens.DOCUMENTCLASS; }
\\usepackage    { BEGIN(pre_optional); return (int) Tokens.USEPACKAGE; }
\\title         { return (int) Tokens.TITLE; }
\\author        { return (int) Tokens.AUTHOR; }
\\today         { return (int) Tokens.TODAY; }
\\date          { return (int) Tokens.DATE; }

// Environments ( + environments specific)
// -----------------------------------------------------------------------------
{envBegin}\{document\}      { return (int) Tokens.BEGIN_DOCUMENT; }
{envEnd}\{document\}        { return (int) Tokens.END_DOCUMENT; }
{envBegin}\{frame\}         { return (int) Tokens.BEGIN_FRAME; }
{envEnd}\{frame\}           { return (int) Tokens.END_FRAME; }
{envBegin}\{itemize\}       { return (int) Tokens.BEGIN_ITEMIZE; }
{envEnd}\{itemize\}         { return (int) Tokens.END_ITEMIZE; }
{envBegin}\{enumerate\}     { return (int) Tokens.BEGIN_ENUMERATE; }
{envEnd}\{enumerate\}       { return (int) Tokens.END_ENUMERATE; }
{envBegin}\{description\}   { return (int) Tokens.BEGIN_DESCRIPTION; }
{envEnd}\{description\}     { return (int) Tokens.END_DESCRIPTION; }
{envBegin}\{tabular\}       { tabular = true; return (int) Tokens.BEGIN_TABULAR; }
{envEnd}\{tabular\}         { tabular = false; return (int) Tokens.END_TABULAR; }
\\item                      { BEGIN(pre_overlay); return (int) Tokens.ITEM; }
\\multicolumn               {}

// Beamer specific
// -----------------------------------------------------------------------------
\\frame         { return (int) Tokens.FRAME; }
\\frametitle    { return (int) Tokens.FRAMETITLE; }
\\framesubtitle { return (int) Tokens.FRAMESUBTITLE; }
\\pause         {}
\\usetheme      { return (int) Tokens.USETHEME; }

// Text formatting
// -----------------------------------------------------------------------------
\\textbf        { return (int) Tokens.TEXTBF; }
\\texttt        { return (int) Tokens.TEXTTT; }
\\textit        { return (int) Tokens.TEXTIT; }
\\textsc        { return (int) Tokens.TEXTSC; }
\\bfseries      { return (int) Tokens.BFSERIES; }
\\ttfamily      { return (int) Tokens.TTFAMILY; }
\\itshape       { return (int) Tokens.ITSHAPE; }
\\scshape       { return (int) Tokens.SCSHAPE; }
\\tiny          { return (int) Tokens.TINY; }
\\scriptsize    { return (int) Tokens.SCRIPTSIZE; }
\\footnotesize  { return (int) Tokens.FOOTNOTESIZE; }
\\small         { return (int) Tokens.SMALL; }
\\normalsize    { return (int) Tokens.NORMALSIZE; }
\\large         { return (int) Tokens.LARGE; }
\\Large         { return (int) Tokens.LARGE2; }
\\LARGE         { return (int) Tokens.LARGE3; }
\\huge          { return (int) Tokens.HUGE; }
\\Huge          { return (int) Tokens.HUGE2; }
\\color         { return (int) Tokens.COLOR; }
\\underline     { return (int) Tokens.UNDERLINE; }

// Images
// -----------------------------------------------------------------------------
\\includegraphics   {}
\\graphicspath      {}

// Other
// -----------------------------------------------------------------------------
\\titlepage     {}
\\hline         {}
\\cline         {}
\\section       { return (int) Tokens.SECTION; }
\\subsection    { return (int) Tokens.SUBSECTION; }
\\subsubsection { return (int) Tokens.SUBSUBSECTION; }
\\LaTeX[ ]?     { BEGIN(str); unformattedText = @"LaTeX"; spaces = 0; nls = 0; }
// New paragraph
\\\\|\\cr       { if(tabular) return (int) Tokens.ENDROW; else return (int) Tokens.NL; }

// Short space, todo copy to plain text loop
\\[ ]           {}
\{              { return '{'; }
\}              { return '}'; }
&               { return '&'; }
// Comments
%.*\r?\n?{ws}    {/*ignore*/}

// White space at the begining of line -> really needed?
^{wsp}         {/*ignore*/}

// Overlay specification
// -----------------------------------------------------------------------------
<pre_overlay> {
        {wsl}\<                 BEGIN(overlay); unformattedText = "";
        [^({wsl}\<)]            BEGIN(pre_optional); yyless(0);
    }
<overlay> {
        [^\>]*                  unformattedText += yytext;
        \>                      BEGIN(pre_optional); yylval.Text = unformattedText; return (int) Tokens.OVERLAY;
    }

// Optional parameters
// -----------------------------------------------------------------------------
<pre_optional> {
        {wsl}\[                 BEGIN(optional); unformattedText = "";
        [^({wsl}\[)]            BEGIN(INITIAL); yyless(0);
    }
<optional> {
        [^\]]*                  unformattedText += yytext;
        \]                      BEGIN(INITIAL); yylval.Text = unformattedText; return (int) Tokens.OPTIONAL;
    }

// Plain text
// -----------------------------------------------------------------------------
[^#\$%\^&_\{\}~\\[:IsWhiteSpace:]]    BEGIN(str); yyless(0); unformattedText = ""; spaces = 0; nls = 0;

<str> {
        [^#\$%\^&_\{\}~\\ \n\t\r]*      { unformattedText += yytext; spaces = 0; nls = 0; }
        " "|\t                     {    // eat multiple spaces
                                        spaces++;
                                        if(spaces == 1)
                                            unformattedText += @" ";
                                   }
        \r                         { /* ignore */ }
        \n                         {    // eat multiple new lines
                                        nls++;
                                        if(nls == 1 && spaces == 0) {   // if one empty line add space
                                            unformattedText += @" ";
                                            spaces++;
                                        }
                                        if(nls == 2) {  // if two empty lines remove space and add new line
                                            // need to remove last space
                                            unformattedText.Remove(unformattedText.Length - 1, 1);
                                            unformattedText += "\n";
                                        }
                                   }
        // escape sequences todo: copy to INITIAL and start str
        \\#                        unformattedText += @"#"; spaces = 0; nls = 0;
        \\\$                       unformattedText += @"$"; spaces = 0; nls = 0;
        \\%                        unformattedText += @"%"; spaces = 0; nls = 0;
        \\textasciicircum\{\}      unformattedText += @"^"; spaces = 0; nls = 0;
        \\&                        unformattedText += @"&"; spaces = 0; nls = 0;
        \\_                        unformattedText += @"_"; spaces = 0; nls = 0;
        \\\{                       unformattedText += @"{"; spaces = 0; nls = 0;
        \\\}                       unformattedText += @"}"; spaces = 0; nls = 0;
        \\~\{\}                    unformattedText += @"~"; spaces = 0; nls = 0;
        \\textbackslash\{\}        unformattedText += @"\"; spaces = 0; nls = 0;
        \\textpipe(\{\})?          unformattedText += @"|"; spaces = 0; nls = 0;
        \\LaTeX[ ]?                unformattedText += @"LaTeX"; spaces = 0; nls = 0; // todo: fix possible parameters
        %.*\n?{ws}                 {/* ignore comment inside plaintext */}
        // end of plain text
        [#\$\^&_\{\}~\\]       BEGIN(INITIAL); yyless(0); yylval.Text = unformattedText; return (int) Tokens.STRING;
    }

\\[[:IsLetter:]]+   {Console.WriteLine("Unknown command!");}

<<EOF>>                            {/* to process, or not to process? */}
%%

override public void yyerror(string format, params object[] args)
{
    string tmp;

    tmp = System.String.Format("{0}:{1} - ", yyline, yycol);
    Messenger.Instance.SendMessage(tmp + format, SimpleConverter.Contract.MessageLevel.ERROR);
}
