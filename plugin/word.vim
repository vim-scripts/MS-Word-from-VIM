" File:                  plugin/word.vim
" Purpose:               To interact with Microsoft Word from VIM
"                        For example, to use Word's Thesaurus and Speller
"
" Copyright (C) 2004 Suresh Govindachar  <initial><last name><at><yahoo>
"
" Version:               1.10
" Date:                  September 9, 2004
" Initial Release:       September 4, 2004
" Documentation:         See below (General Notes) 
"
" General Notes:                                           {{{1
"
" System Requirements:
"   
"    Windows (with MS Word)
"    VIM with embedded perl support
"    The perl module Win32::OLE (good way to get this is to have external perl)
" 
" Usage:  
"
"    The following four commands are exported to the end-user:
"
"       :WordThesaurus    [<cword>]
"       :WordSpell        [<cword>]
"       :WordAnagram      [<cword>]
"       :<range>WordCheckSpelling 
"
"    In regard to the first three commands:
"
"     - Each command can be issued with no arguments or with one argument.
"       When no argument is provided, the <cword> (see :help cword) is 
"       used as the argument.
"     - After execution, the "result" is appended below the current line.
"     - The "result" for the Thesaurus command is all 
"       the synonyms and the antonyms.
"     - The "result" for the Spell command is all 
"       the spelling suggestions.
"     - The "result" for the Anagram command is all 
"       the anagrams of the given letters.
"
"    Try the following:
"
"       :WordThesaurus beautiful
"       :WordAnagram   abt
"       :WordSpell     helo
"       :WordSpell     b?t*
"
"    In regard to the fourth command (:<range>WordCheckSpelling):
"       
"     - The default range is the whole buffer -- in order for
"       quickfix commands to work, there has to be a file 
"       associated with the buffer.
"
"     - Results are:
"
"       - highlighting as Error each misspelled word in the range
"         (the filetype is set to nothing so that the only colored
"         stuff are the misspelled words)
"
"       - ability to use the quickfix commands (:cfirst, :cn, etc) to 
"         visit each misspelled word and to look at suggestions on spelling it.
"
"       - Sub commands (of the :<range>WordCheckSpelling command):
"
"            :WordShowMisspelled
"                 Highlights misspelled words (there won't be any other highlighting)                 
"            :WordHideMisspelled
"                 Restores highlighting based on filetype 
"            The next five commands extend the corresponding quickfix 
"                 commands (:cfirst, :clast, :cc, :cn, :cp) by
"                 loading the search register with the appropriate 
"                 misspelled word
"                     :Wordfirst
"                     :Wordlast
"                     :Wordc
"                     :Wordn
"                     :Wordp
"
"       - Optional (controlled by g:no_plugin_maps or g:no_word_maps) 
"         normal-mode maps:
"            <Leader>ws :WordShowMisspelled<CR>
"            <Leader>wh :WordHideMisspelled<CR>
"            <Leader>wc :Wordc<CR>
"            <Leader>wn :Wordn<CR>
"            <Leader>wp :Wordp<CR>
"
"       - Non-user level effects:
"
"          - On sourcing this file, 'set errorformat+=%f:%l(%c):%m' is
"            executed
"          - The quickfix tags file for WordCheckSpelling is 
"            called word_misspelled.tags. and is in the same directory 
"            as VIM's tempname().
"          - The variable g:word_misspelled will be a space separated 
"            string of all the misspelled words.
"
"    Tested on:
"
"       Windows 98 with Word 97 (and Vim 6.3, ActiveState's perl 5.8).
"
"    Known issue:
"      
"       Sometimes Windows and/or Word is sluggish and things
"       time-out with a "no service available" or some such 
"       error message -- when this happens, just reissue the command.
"
"----------------------------------------------------------------------------- {{{2
"       I thought of implementing the following interactive commands
"       but decided against doing so because of time.
"       So, the following interactive commands have NOT been implemented
"            :WordThesaurusInteractive   [<cword>]
"            :WordSpellInteractive       [<cword>]
"
"          - The Interactive commands allow the user to interact 
"            with Word's Thesaurus or Spelling dialog box and to make a
"            selection there-after there-in.
"          - The "result" for the Interactive commands is the user's choice.
"
"----------------------------------------------------------------------------- 2}}}
"
" Acknowledgment:                                                              {{{1
"
"          Although I conceived and developed this plugin, I
"          learnt how to interact with Microsoft Word for 
"          accessing its Thesaurus and Spelling from: 
"
"           - Chad DeMeyer   (microsoft.public.word.word97vba)
"           - Steven Manross (perl-win32-users@listserv.ActiveState.com)
"           - Greg Chapman   (perl-win32-users@listserv.ActiveState.com)
"           - http://www.perlmonks.com/index.pl?node_id=119006
"
"---------------------------------------------------------------------------
"
" Disclaimer:                                                                 {{{1
"
"     The material provided here is provided as-is without any warranty --
"     without even the implied warranty of merchantability or fitness for a
"     particular purpose.  The author assumes no responsibility for errors or
"     omissions or for any sort of damages resulting from the use of or
"     reliance on the provided material.
"
"
"bookkeeping {{{1
if exists("loaded_word")  "{{{2
   finish
endif
if !has('perl')  
   let g:loaded_word = 'no perl'
   finish
endif
if !has("win32")
   let g:loaded_word = 'not on windows'
   finish
endif
let g:loaded_word = 'pre test of perl version'
perl << EOVersionTest
   require 5.8.0;
   VIM::DoCommand('let g:loaded_word=\'passed test of perl version\'');
EOVersionTest
if (g:loaded_word != 'passed test of perl version')
   let g:loaded_word = 'failed test of perl version'
   finish
endif

if !exists("word_language")
   let g:word_language = 0
endif
"let word_debug = 1
if !exists("word_debug")
   let g:word_debug = 0
endif
let g:word_misspelled=' '

"see help use-cpo-save for info on the variable save_cpo  
let s:save_cpo = &cpo
set cpo&vim

"perl  {{{1
"
"Core {{{2
"
perl << EOCore

#BEGIN {(*STDERR = *STDOUT) || die;} # {{{4
#line 77
# put=line(\".\")
# normal u
# dis =
# normal zR
# %g/^#line .*/ exec 'normal 6lD'|put =|normal kJ

use diagnostics;
use warnings;

use strict;
use Win32::OLE;
use Win32::OLE::Const 'Microsoft Word';

#   VIM::Msg("Compiling core ..."); 

   my $debug         = VIM::Eval('g:word_debug');       # {{{3
   my $word_language = VIM::Eval('g:word_language'); 
      $word_language = wdEnglishUS unless $word_language;

   my $tags_file     = VIM::Eval('fnamemodify(tempname(), ":p:h")') . "/word_misspelled.tags"; 
                       VIM::DoCommand('set errorformat+=%f:%l(%c):%m');

# word_look_up_as($as_what, $word, $maybe) # {{{3
sub word_look_up_as # {{{4
{
   my ($as_what, $word) = @_;  # see how default <cword> just happens! 

   $word =~ m/^[a-zA-Z*?]+$/ or return "\n"; # only allow safe characters # no hyphenated words

   my $wildness = ($word =~ /\W/) ? wdWildcard : wdSpellword;
   $wildness = wdAnagram if ($as_what =~ /anagram/i); 

   my ($word_obj, $document_obj) = word_initialize($word);
   my $result = '';
    
   if($as_what =~ /thesaurus/i)
   {
                # cannot feed wildness to thesaurus!!!
      $result = word_thesaurus($word_obj, $word, $word_language) unless ($word =~ m/\W/);
   }
   else 
   {
                # cannot feed language to spelling!!!
      $result = word_spelling($word_obj, $word, $wildness); 
   }
   
   word_shut_down($document_obj, $word_obj);
    
###word_echo('msg', $result);
$result = " \n$result\n ";
my @ar = split /\n/, $result;
$main::curbuf->Append(($main::curwin->Cursor())[0], @ar);

   return $result;
}

# word_check_spelling(<line1>, <line2>) # {{{3
#
# The command first picks out unique sequences of 
# contiguous alphabets ([a-zA-Z]+) ("words") and records 
# them along with their location information.  It then issues
# the WordSpell command on each word.  If Microsoft Word provides 
# any suggestions for spelling it then this word together with
# the suggestions and its recorded location are 
# recorded and later processed by VIM's quickfix and 
# highlighter (with the Error highlight).
#
# To turn highlighting on and off, the misspelled words
# are also stored as a space separated string g;word_misspelled
#
sub word_check_spelling # {{{4
{
  my ($line, $line_end) = @_;
  #my @lines  = $main::curbuf->Get(1 .. $main::curbuf->Count());
  my @lines  = $main::curbuf->Get($line .. $line_end);
  my $prefix = VIM::Eval('expand("%:p")') . ':'; 

  my $cursor   = 0;
  my $position = $cursor;
  my $word     = '';

  my %all_words=();
  my @ordered_words=();
  #my @ordered_prefixes=();
  for (@lines)
  {
     $cursor = 1; 

     while (m/([^a-zA-Z]*([a-zA-Z]+)[^a-zA-Z]*)/g) #while (m/([^a-zA-Z]*([a-zA-Z]+-?[a-zA-Z]*)[^a-zA-Z]*)/g)
     {
        $position     = $cursor;
        $cursor      += length($1);
        $word         = $2;

        exists $all_words{$word} and next;
        $all_words{$word} = "$prefix$line($position):$word -- ";
        push @ordered_words,  $word;
        #push @ordered_prefixes, "$prefix$line($position):$word -- ";
     } 
     $line++;
  }
  my $misspelt = ' '; # a misspelling of misspelled!
  my @tags     = ();
  
  my $result = '';
  my $wildness = wdSpellword;
  my ($word_obj, $document_obj) = word_initialize('spell');
  
  #my $i = 0;
  #my $foo = '';
  for $word (@ordered_words)
  {
     #$foo    = $ordered_prefixes[$i++];
             # cannot feed language to spelling!!!
     $result = word_spelling($word_obj, $word, $wildness); 
     $result =~ s/^\s*spelling\s+//i;
     $result =~ /\w/ or next;
  
     $misspelt .= "$word ";
     push @tags, $all_words{$word} . $result;
     #push @tags, $foo . $result;
  }
  
  word_shut_down($document_obj, $word_obj);
  
  VIM::DoCommand("let g:word_misspelled = \'$misspelt\'");
  if(@tags)
  {
     VIM::DoCommand('set ft= ');
     VIM::DoCommand('syn clear WordMisspelled');
     VIM::DoCommand('syn keyword WordMisspelled '. $misspelt);
     VIM::DoCommand('hi link WordMisspelled Error');
  
     if(open(TAGS, ">$tags_file")) # or die "Unable to open $tags_file for writing:$!\n"; 
     {
        print TAGS @tags;
        close TAGS;
        VIM::DoCommand("cfile $tags_file");
     }
  }
  else
  {
     VIM::Msg('No misspelled words!', 'comment'); 
  }
}

# word_initialize($word)  # {{{3
sub word_initialize  # {{{4
{
  (my $word_obj = Win32::OLE->new("Word.Application"))
               or die "No word object:\n$!\n".Win32::OLE->LastError()."\n";

  # thesaurus does not need a document object, 
  # but spellings needs it to be created!!!
  (my $document_obj = $word_obj->Documents->Add())
               or warn "No document object:\n$!\n".Win32::OLE->LastError()."\n"
               and word_shut_down(0, $word_obj);

   $word_obj->{WindowState} = 0;
   $document_obj->Activate();
   #$document_obj->Range->InsertBefore($word);

   return $word_obj;
}

# $result = word_thesaurus($word_obj, $word, $word_language) # {{{3
sub word_thesaurus # {{{4
{
  my ($word_obj, $word, $language) = @_;
  my $result = "Synonyms:\n";

  # no need to have created document_obj 

  my $syninfo = $word_obj->SynonymInfo({Word=>$word, LanguageID=>$language});

     $syninfo or  warn "No syninfo object:\n$!\n".Win32::OLE->LastError()."\n"
              and return "\n";

     ($word eq ${$syninfo}{Word}) or return "MANGLED\n"; 
     $syninfo->Found or return "\n";

  my $matches = $syninfo->MeaningCount;
  for(1 .. $matches)
  {
     foreach (@{$syninfo->SynonymList($_)}) 
     {
        $result .= "$_  ";
     }
     $result .= "\n";
  }

  my $antonyms=$syninfo->AntonymList; 
  if($antonyms) 
  {
     $result .= "Antonyms:\n";
     foreach (@{$antonyms})
     {
        $result .= "$_  ";
     }
     $result .= "\n";
  }

  return $result;
}

# $result = word_spelling($word_obj, $word, $wildness) # {{{3
sub word_spelling # {{{4
{
  my ($word_obj, $word, $wildness) = @_;
  my $result = "Spelling\n";

  # eventhough document_obj is not explicitly used, it must be created!!!

  my $spellinfo = $word_obj->GetSpellingSuggestions({Word=>$word, SuggestionMode=>$wildness});
     $spellinfo or  warn "No spellinfo object:\n$!\n".Win32::OLE->LastError()."\n"
                and return "\n";

  my $count = $spellinfo->Count; 
     $count or return "\n";

  for(1 .. $count)
  {
     $result .= $spellinfo->Item($_)->Name ."  ";
  }
  $result .= "\n";
  return $result;
}

# word_shut_down($document_obj, $word_obj) # {{{3
sub word_shut_down # {{{4
{
  my ($document_obj, $word_obj) = @_;

  if($document_obj) 
  {
    $document_obj->{Saved} = 0;
    $document_obj->Close(0);
  }
  if($word_obj) 
  {
    $word_obj->{Visible} = 0;
    $word_obj->Quit();
  }
}

# word_debug($what) # {{{3
sub word_debug # {{{4
{
  my $what = shift;
  $debug and word_echo('msghlsearch', $what);
}

# word_echo($type, $what) # {{{3
# $what is a multi-line ("\n") string that will be echo'ed
# with the highlighting encoded in $type
#
# 2do: escape things in $what which VIM cannot echo.
# word_echo # {{{3
sub word_echo # {{{4
{
  my ($type, $what) = @_;
  $type = 'echo'.$type;        #$type can be msg or msgwarn or err or msghlToDo etc.

  ($type =~ s/warn//)   and VIM::DoCommand('echohl WarningMsg'); 
  ($type =~ s/hl(.+)//) and VIM::DoCommand("echohl $1"); 
  foreach (split "\n", $what)
  {
     VIM::DoCommand("$type \'$_\'");
  }
  VIM::DoCommand('echohl None');
  return 1;
}

# word_die($dying_message) # {{{3
sub word_die # {{{4
{
  my ($dying_message) = @_;
  word_echo('err', $dying_message);
  die $dying_message;
}
   
word_debug("                   1 Done compiling core\n");

EOCore

" Setting up the commands: {{{1
" Recall                   {{{2
"       :WordThesaurus    [<cword>]
"       :WordSpell        [<cword>]
"       :WordAnagram      [<cword>]
"
"perl word_look_up_as('thesaurus', 'beautiful');
"perl word_look_up_as('anagram', 'bat');
"perl word_look_up_as('spelling', 'b?t');

"help user-commands    {{{2
" the default of <cword> will just happen by magic during reading! 
command! -nargs=?          WordThesaurus     :perl word_look_up_as('thesaurus', <f-args>, scalar VIM::Eval('expand("<cword>")')); 
command! -nargs=?          WordSpell         :perl word_look_up_as('spelling',  <f-args>, scalar VIM::Eval('expand("<cword>")')); 
command! -nargs=?          WordAnagram       :perl word_look_up_as('anagram',   <f-args>, scalar VIM::Eval('expand("<cword>")')); 
command! -nargs=0 -range=% WordCheckSpelling :perl word_check_spelling(<line1>, <line2>)
command! WordShowMisspelled :if(word_misspelled =~ '\S') | set ft= | let word_foo=&iskeyword | set iskeyword=a-z,A-Z | exec 'syn keyword WordMisspelled '. word_misspelled | hi link WordMisspelled error | redraw | execute 'set iskeyword='. word_foo | endif
command! WordHideMisspelled :hi link WordMisspelled none | filetype detect
command! Wordfirst        :cfirst | exe "let @/ ='" . expand("<cword>"). "'"
command! Wordlast         :clast  | exe "let @/ ='" . expand("<cword>"). "'"
command! Wordc            :cc | exe "let @/ ='" . expand("<cword>"). "'"
command! Wordn            :cn | exe "let @/ ='" . expand("<cword>"). "'"
command! Wordp            :cp | exe "let @/ ='" . expand("<cword>"). "'"

if ((!exists("no_plugin_maps")) && (!exists("no_word_maps")))
   nmap <Leader>wc :Wordc<CR>
   nmap <Leader>wn :Wordn<CR>
   nmap <Leader>wp :Wordp<CR>
   nmap <Leader>ws :WordShowMisspelled<CR>
   nmap <Leader>wh :WordHideMisspelled<CR>
endif

"restore saved cpo     {{{1
let &cpo = s:save_cpo

"<SID> any functions?

finish

