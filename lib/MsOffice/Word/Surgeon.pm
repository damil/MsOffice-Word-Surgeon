package MsOffice::Word::Surgeon;
use 5.010;
use Moose;
use Archive::Zip                          qw(AZ_OK);
use Encode;
use List::Util                            qw(pairs);
use List::MoreUtils                       qw(uniq);
use XML::LibXML;
use POSIX                                 qw(strftime);
use MsOffice::Word::Surgeon::Utils        qw(maybe_preserve_spaces);
use MsOffice::Word::Surgeon::Replacement;
use namespace::clean -except => 'meta';

# constant integers to specify indentation modes -- see L<XML::LibXML>
use constant XML_NO_INDENT     => 0;
use constant XML_SIMPLE_INDENT => 1;

# name of the zip member that contains the main document body
use constant MAIN_DOCUMENT => 'word/document.xml';


has 'filename'    => (is => 'ro', isa => 'Str', required => 1);

has 'zip'         => (is => 'ro',   isa => 'Archive::Zip',
                      builder => '_zip',   lazy => 1);

has 'contents'    => (is => 'rw',   isa => 'Str',
                      builder => 'original_contents', lazy => 1);

has 'last_rev_id' => (is => 'ro',   isa => 'Num', default => 0);

# infos to be used in revision marks
has 'author'      => (is => 'ro',   isa => 'Str', default => 'Surgeon');
has 'date'        => (is => 'ro',   isa => 'Str',
                      default => sub {strftime "%Y-%m-%dT%H:%M:%SZ", localtime});



#======================================================================
# GLOBAL VARIABLES
#======================================================================

# Various regexes for removing XML information without interest
my %noise_reduction_regexes = (
  proof_checking      => qr(<w:(?:proofErr[^>]+|noProof/)>),
  revision_ids        => qr(\sw:rsid\w+="[^"]+"),
  complex_script_bold => qr(<w:bCs/>),
  page_breaks         => qr(<w:lastRenderedPageBreak/>),
  language            => qr(<w:lang w:val="[^"]+"/>),
  empty_para_props    => qr(<w:rPr></w:rPr>),
 );

my @noise_reduction_list = qw/proof_checking  revision_ids complex_script_bold
                              page_breaks language empty_para_props/;

# Regex to split the whole contents into OOXML "runs" - and fragments inbetween
# See $self->merge_runs()
my $run_regex = qr[
   <w:r>                             # opening tag for the run
   (?:<w:rPr>(.*?)</w:rPr>)?         # run properties -- capture in $1
   <w:t(?:\ xml:space="preserve")?>  # opening tag for the text contents
   (.*?)                             # text contents -- capture in $2
   </w:t>                            # closing tag for text
   </w:r>                            # closing tag for the run
  ]x;

$run_regex = qr[
   <w:r>                             # opening tag for the run
   (?:<w:rPr>(.*?)</w:rPr>)?         # run properties -- capture in $1
   (?:
     <w:t(?:\ xml:space="preserve")?>  # opening tag for the text contents
     (.*?)                             # text contents -- capture in $2
     </w:t>                            # closing tag for text
   )?
   </w:r>                            # closing tag for the run
  ]x;





# regexes for unlinking MsWord fields
my $field_instruction_txt_rx = qr[<w:instrText.*?</w:instrText>];
my $field_boundary_rx        = qr[<w:fldChar.*?/>]; #  "begin" / "separate" / "end"
my $simple_field_rx          = qr[</?w:fldSimple[^>]*>];
my $empty_run_rx             = qr[<w:r>(?:<w:rPr>.*?</w:rPr>)?</w:r>];

#======================================================================
# BUILDING
#======================================================================


# syntactic sugar for ->new($path) instead of ->new(filename => $path)
around BUILDARGS => sub {
  my $orig  = shift;
  my $class = shift;

  if ( @_ == 1 && !ref $_[0] ) {
    return $class->$orig(filename => $_[0]);
  }
  else {
    return $class->$orig(@_);
  }
};



#======================================================================
# LAZY ATTRIBUTE CONSTRUCTORS
#======================================================================

sub _zip {
  my $self = shift;

  my $zip = Archive::Zip->new;
  $zip->read($self->{filename}) == AZ_OK
      or die "cannot unzip $self->{filename}";

  return $zip;
}

sub original_contents { # can also be called later, not only as lazy constructor
  my $self = shift;

  my $bytes    = $self->zip->contents(MAIN_DOCUMENT)
    or die "no contents for member ", MAIN_DOCUMENT();
  my $contents = decode_utf8($bytes);
  return $contents;
}


#======================================================================
# CONTENTS RESTITUTION
#======================================================================

sub indented_contents {
  my $self = shift;

  my $dom = XML::LibXML->load_xml(string => $self->contents);
  return $dom->toString(XML_SIMPLE_INDENT); # returned as bytes sequence, not a Perl string
}



sub plain_text {
  my $self = shift;

  # XML contents
  my $txt = $self->contents;

  # replace opening paragraph tags by newlines
  $txt =~ s/(<w:p[ >])/\n$1/g;

  # remove all remaining XML tags
  $txt =~ s/<[^>]+>//g;

  return $txt;
}


#======================================================================
# MODIFYING CONTENTS
#======================================================================

sub reduce_noise {
  my ($self, @noises) = @_;

  my $contents = $self->contents;

  my @regexes = map {ref $_ eq 'Regexp' ? $_ : $noise_reduction_regexes{$_}} @noises;
  $contents =~ s/$_//g foreach @regexes;

  $self->contents($contents);
}

sub reduce_all_noises {
  my $self = shift;

  $self->reduce_noise(@noise_reduction_list);
}


sub merge_runs {
  my $self = shift;

  # split the whole contents on run boundaries
  my $contents  = $self->contents;
  my @fragments = split m[$run_regex], $contents, -1;

  # variables to iterate on fragments and reconstruct the contents
  my @new_fragments;
  my $last_props;
  my $last_t_contents;

  # loop on triplets in @fragments
  while (my ($before_run, $run_props, $t_contents) = splice @fragments, 0, 3) {
    no warnings 'uninitialized'; # because some @fragments may be undef

    if (!$before_run && $run_props eq $last_props) {
      # merge this run contents with the previous run
      $last_t_contents .= $t_contents;
    }
    else {
      if ($last_t_contents) {
        # emit previous run
        my $props             = $last_props ? "<w:rPr>$last_props</w:rPr>" : ""; 
        my $space_attr        = maybe_preserve_spaces($last_t_contents);
        push @new_fragments, "<w:r>$props<w:t$space_attr>$last_t_contents</w:t></w:r>";
      }

      # emit contents preceding the current run
      push @new_fragments, $before_run if $before_run;

      # current run becomes "previous run"
      $last_props      = $run_props;
      $last_t_contents = $t_contents;
    }
  }

  # reassemble the whole stuff and inject it as new contents
  $self->contents(join "", @new_fragments);
}


sub unlink_fields {
  my $self = shift;

  # Note : fields can be nested, so their internal structure is quite complex
  # ... but after various unsuccessful attempts with grammars I finally found a
  # much easier solution
  $self->reduce_noise($field_instruction_txt_rx, $field_boundary_rx,
                      $simple_field_rx); ###, $empty_run_rx);
}



sub apply_replacements {
  my ($self, @replacements) = @_; # list of pairs [$old => $new] -- order will be preserved

  # build a regex of all $old texts to replace
  my @patterns = map {$_->[0]} @replacements;
  $_ =~ s/(\p{Pattern_Syntax})/\\$1/g foreach @patterns;  # escape regex chars
  my $all_patterns = join "|", @patterns;

  # build a substitution callback
  my %replacement = map {@$_} @replacements;
  my $replace_it   = sub {
    my $orig  = shift;
    my $repl = $replacement{$orig};
    my $auth = $repl =~ /^\p{Lu}___/ ? $repl : __PACKAGE__;  # todo : better default
    my $replacer = MsOffice::Word::Surgeon::Replacement->new(
      original    => $orig,
      replacement => $repl,
      author      => $auth,
     );
    return $replacer->as_xml;
  };


  # global substitution, inserting revision marks for each pattern found
  my $contents  = $self->contents;
  $contents =~ s/($all_patterns)/$replace_it->($1)/eg;

  # inject as new contents
  $self->contents($contents);
}





#======================================================================
# SAVING THE FILE
#======================================================================


sub _update_zip {
  my $self = shift;

  $self->zip->contents(MAIN_DOCUMENT, encode_utf8($self->contents));
}


sub overwrite {
  my $self = shift;

  $self->_update_zip;
  $self->zip->overwrite;
}



sub save_as {
  my ($self, $filename) = @_;

  $self->_update_zip;
  $self->zip->writeToFileNamed($filename) == AZ_OK
    or die "error writing zip archive to $filename";
}




1;

__END__

TODO
  - doc
  - tests
  - Surgeon->new($filename) -- find a way in Moose




=head1 SEE ALSO

Here are some packages in other languages that deal with C<docx> documents.

=over

=item L<https://phpword.readthedocs.io/en/latest/>

=item L<https://www.docx4java.org/trac/docx4j>

=item L<https://pypi.org/project/python-docx/>

=item L<https://docs.microsoft.com/en-us/office/open-xml/word-processing>



=back

https://metacpan.org/pod/Document::OOXML::Document::Wordprocessor

https://www.toptal.com/xml/an-informal-introduction-to-docx

