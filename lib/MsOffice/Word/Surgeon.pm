package MsOffice::Word::Surgeon;
use 5.010;
use Moose;
use Archive::Zip                          qw(AZ_OK);
use Encode                                qw(encode_utf8 decode_utf8);
use XML::LibXML;
use MsOffice::Word::Surgeon::Run;
use MsOffice::Word::Surgeon::Text;
use MsOffice::Word::Surgeon::Change;

use namespace::clean -except => 'meta';

# constant integers to specify indentation modes -- see L<XML::LibXML>
use constant XML_NO_INDENT     => 0;
use constant XML_SIMPLE_INDENT => 1;

# name of the zip member that contains the main document body
use constant MAIN_DOCUMENT => 'word/document.xml';


has 'filename'    => (is => 'ro', isa => 'Str', required => 1);

has 'zip'         => (is => 'ro',   isa => 'Archive::Zip',
                      builder => '_zip',   lazy => 1);

has 'contents'    => (is => 'rw',   isa => 'Str', init_arg => undef,
                      builder => 'original_contents', lazy => 1,
                      trigger => sub {shift->clear_runs},
                     );

has 'runs'        => (is => 'ro',   isa => 'ArrayRef', init_arg => undef,
                      builder => '_runs', lazy => 1, clearer => 'clear_runs');



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


# regexes for unlinking MsWord fields
my $field_instruction_txt_rx = qr[<w:instrText.*?</w:instrText>];
my $field_boundary_rx        = qr[<w:fldChar.*?/>]; #  "begin" / "separate" / "end"
my $simple_field_rx          = qr[</?w:fldSimple[^>]*>];

#======================================================================
# BUILDING
#======================================================================


# syntactic sugar for supporting ->new($path) instead of ->new(filename => $path)
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


sub _runs {
  my $self = shift;

  state $run_regex = qr[
    <w:r>                             # opening tag for the run
    (?:<w:rPr>(.*?)</w:rPr>)?         # run properties -- capture in $1
    (.*?)                             # run contents -- capture in $2
    </w:r>                            # closing tag for the run
  ]x;

  state $txt_regex = qr[
    <w:t(?:\ xml:space="preserve")?>  # opening tag for the text contents
    (.*?)                             # text contents -- capture in $1
    </w:t>                            # closing tag for text
  ]x;


  my $contents      = $self->contents;
  my @run_fragments = split m[$run_regex], $contents, -1;
  my @runs;

  while (my ($xml_before_run, $props, $run_contents) = splice @run_fragments, 0, 3) {

    $run_contents //= '';

    my @txt_fragments = split m[$txt_regex], $run_contents, -1;
    my @texts;
    while (my ($bt, $txt_contents) = splice @txt_fragments, 0, 2) {
      push @texts, MsOffice::Word::Surgeon::Text->new(
        xml_before   => $bt           // '',
        literal_text => $txt_contents // '',
       );
    }

    push @runs, MsOffice::Word::Surgeon::Run->new(
      xml_before  => $xml_before_run // '',
      props       => $props          // '',
      inner_texts => \@texts,
     );
  }

  return \@runs;
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

  my @new_runs;
  foreach my $run (@{$self->runs}) {
    if (!$run->xml_before && @new_runs && $new_runs[-1]->props eq $run->props) {
      $new_runs[-1]->merge($run);
    }
    else {
      push @new_runs, $run;
    }
  }

  # reassemble the whole stuff and inject it as new contents
  $self->contents(join "", map {$_->as_xml} @new_runs);
}



sub unlink_fields {
  my $self = shift;

  # Note : fields can be nested, so their internal structure is quite complex
  # ... but after various unsuccessful attempts with grammars I finally found a
  # much easier solution
  $self->reduce_noise($field_instruction_txt_rx, $field_boundary_rx, $simple_field_rx);
}



#======================================================================
# DELEGATION TO SUBCLASSES
#======================================================================

sub change {
  my $self = shift;

  my $change = MsOffice::Word::Surgeon::Change->new(@_);
  return $change->as_xml;
}




#======================================================================
# SAVING THE FILE
#======================================================================


sub _update_contents_in_zip {
  my $self = shift;

  $self->zip->contents(MAIN_DOCUMENT, encode_utf8($self->contents));
}


sub overwrite {
  my $self = shift;

  $self->_update_contents_in_zip;
  $self->zip->overwrite;
}



sub save_as {
  my ($self, $filename) = @_;

  $self->_update_contents_in_zip;
  $self->zip->writeToFileNamed($filename) == AZ_OK
    or die "error writing zip archive to $filename";
}



sub replace {
  my ($self, $pattern, $replacement, @context) = @_;

  my $xml = join "", map {$_->replace($pattern, $replacement, @context)}
                         @{$self->runs};
  return $xml;
}



1;

__END__





TODO
  - doc
  - tests


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

