use strict;
use warnings;
use lib "../lib";
use MsOffice::Word::Surgeon;

(my $dir = $0) =~ s[tst_template.t$][];
$dir ||= ".";
my $template_file = "$dir/etc/tst_template.docx";


my $surgeon = MsOffice::Word::Surgeon->new($template_file);

$surgeon->reduce_all_noises;
$surgeon->merge_runs;


my $template = $surgeon->compile_template(
  highlights => 'yellow',
 );
my %data = (
  foo => 'FOFOLLE',
  bar => 'WHISKY',
  list => [ {name => 'toto', value => 123},
            {name => 'blublu', value => 456},
            {name => 'zorb', value => 987},
           ],
);
my $new_doc = $template->process(data => \%data);
$new_doc->save_as("template_result.docx");


