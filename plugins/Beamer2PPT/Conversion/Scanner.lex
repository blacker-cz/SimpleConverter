%x str, overlay, optional, pre_overlay, pre_optional, tabular_arg

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
    int tbl = 0;
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
{envBegin}\{tabular\}\{     { tabular = true; BEGIN(tabular_arg); tbl = 0; unformattedText = ""; return (int) Tokens.BEGIN_TABULAR; }
{envEnd}\{tabular\}         { tabular = false; return (int) Tokens.END_TABULAR; }
\\item                      { BEGIN(pre_overlay); return (int) Tokens.ITEM; }
\\multicolumn               { return (int) Tokens.MULTICOLUMN; }

// Beamer specific
// -----------------------------------------------------------------------------
\\frame         { return (int) Tokens.FRAME; }
\\frametitle    { BEGIN(pre_overlay); return (int) Tokens.FRAMETITLE; }
\\framesubtitle { BEGIN(pre_overlay); return (int) Tokens.FRAMESUBTITLE; }
\\pause         { return (int) Tokens.PAUSE; }
\\usetheme      { return (int) Tokens.USETHEME; }

// Text formatting
// -----------------------------------------------------------------------------
\\textbf        { BEGIN(pre_overlay); return (int) Tokens.TEXTBF; }
\\texttt        { BEGIN(pre_overlay); return (int) Tokens.TEXTTT; }
\\textit        { BEGIN(pre_overlay); return (int) Tokens.TEXTIT; }
\\textsc        { BEGIN(pre_overlay); return (int) Tokens.TEXTSC; }
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
\\color         { BEGIN(pre_overlay); return (int) Tokens.COLOR; }
\\textcolor     { BEGIN(pre_overlay); return (int) Tokens.TEXTCOLOR; }
\\underline     { BEGIN(pre_overlay); return (int) Tokens.UNDERLINE; }
\\and           { return (int) Tokens.AND; }

// Images
// -----------------------------------------------------------------------------
\\includegraphics   {}
\\graphicspath      {}

// Other
// -----------------------------------------------------------------------------
\\titlepage     {}
\\hline         { return (int) Tokens.HLINE; }
\\cline         { return (int) Tokens.CLINE; }
\\section       { return (int) Tokens.SECTION; }
\\subsection    { return (int) Tokens.SUBSECTION; }
\\subsubsection { return (int) Tokens.SUBSUBSECTION; }
\\LaTeX[ ]?     { BEGIN(str); unformattedText = @"LaTeX"; spaces = 0; nls = 0; }
// New paragraph
\\\\|\\cr       { BEGIN(pre_optional); if(tabular) return (int) Tokens.ENDROW; else return (int) Tokens.NL; }

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
        {wsl}<                  BEGIN(overlay); unformattedText = "";
        {wsl}[^<]               BEGIN(pre_optional); yyless(0);
    }
<overlay> {
        [^>]*                   unformattedText += yytext;
        >                       BEGIN(pre_optional); yylval.Text = unformattedText; return (int) Tokens.OVERLAY;
    }

// Optional parameters
// -----------------------------------------------------------------------------
<pre_optional> {
        {wsl}\[                 BEGIN(optional); unformattedText = "";
        {wsl}[^[]               BEGIN(INITIAL); yyless(0);
    }
<optional> {
        [^\]]*                  unformattedText += yytext;
        \]                      BEGIN(INITIAL); yylval.Text = unformattedText; return (int) Tokens.OPTIONAL;
    }

// Tabular settings
// -----------------------------------------------------------------------------
<tabular_arg> {
        \{                      tbl++; unformattedText += @"{";
        \}                      {
                                    tbl--; 
                                    if(tbl < 0) {
                                        yylval.Text = unformattedText;
                                        BEGIN(INITIAL);
                                        return (int) Tokens.STRING;
                                    }
                                    unformattedText += @"}";
                                }
        [^{}]*                  unformattedText += yytext;
}

// Plain text
// -----------------------------------------------------------------------------
[^#\$%\^&_\{\}~\\[:IsWhiteSpace:]]    BEGIN(str); yyless(0); unformattedText = ""; /*spaces = 0;*/ nls = 0;

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

// unknown commands etc.
{envBegin}\{[^\}]+\}   { printWarning("Unknown environment " + yytext); }
{envEnd}\{[^\}]+\}     { printWarning("Unknown environment " + yytext); }

\\[[:IsLetter:]]+      { printWarning("Unknown command " + yytext); }

<<EOF>>                            {/* to process, or not to process? */}
%%

override public void yyerror(string format, params object[] args)
{
    string tmp;

    tmp = System.String.Format("{0}:{1} - ", yyline, yycol);
    Messenger.Instance.SendMessage(tmp + format, SimpleConverter.Contract.MessageLevel.ERROR);
}

/// <summary>
/// Print warning message
/// </summary>
/// <param name="message">Message</param>
private void printWarning(string message)
{
    string tmp;

    tmp = System.String.Format("{0}:{1} - ", yyline, yycol);
    Messenger.Instance.SendMessage(tmp + message, SimpleConverter.Contract.MessageLevel.WARNING);
}
