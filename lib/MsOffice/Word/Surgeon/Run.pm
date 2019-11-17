package MsOffice::Word::Surgeon::Run;
use feature 'state';
use Moose;
use Carp qw(croak);

use namespace::clean -except => 'meta';

has 'xml_before'  => (is => 'ro', isa => 'Str', required => 1);
has 'props'       => (is => 'ro', isa => 'Str', required => 1);
has 'inner_texts' => (is => 'ro', required => 1,
                      isa => 'ArrayRef[MsOffice::Word::Surgeon::Text]');


sub as_xml {
  my $self = shift;
  my $xml  = $self->xml_before;
  if (@{$self->inner_texts}) {
    $xml .= "<w:r>";
    $xml .= "<w:rPr>" . $self->props . "</w:rPr>" if $self->props;
    $xml .= $_->as_xml foreach @{$self->inner_texts};
    $xml .= "</w:r>";
  }

  return $xml;
}



sub merge {
  my ($self, $next_run) = @_;

  # sanity checks
  $next_run->isa(__PACKAGE__)
    or croak "argument to merge() should be a " . __PACKAGE__;
  $self->props eq $next_run->props
    or croak sprintf "runs have different properties: '%s' <> '%s'",
                      $self->props, $next_run->props;
  !$next_run->xml_before
    or croak "cannot merge -- next run contains xml before the run : "
           . $next_run->xml_before;

  # loop over all text nodes of the next run
  foreach my $txt (@{$next_run->inner_texts}) {
    if (@{$self->{inner_texts}} && !$txt->xml_before) {
      # concatenate current literal text with the previous text node
      $self->{inner_texts}[-1]->merge($txt);
    }
    else {
      # cannot merge, just add to the list of inner text nodes
      push @{$self->{inner_texts}}, $txt;
    }
  }
}


sub replace {
  my ($self, $pattern, $replacement, %args) = @_;

  my @inner_xmls 
    = map {$_->replace($pattern, $replacement, run => $self, %args)}
          @{$self->inner_texts};

  my $xml = $self->xml_before . join "", @inner_xmls;

  return $xml;
}



1;

__END__

=encoding ISO-8859-1

=head1 NAME

MsOffice::Word::Surgeon::Run -- internal representation for a "run of text"

=head1 DESCRIPTION

This is used internally by L<MsOffice::Word::Surgeon> for storing
a "run of text" in and MsWord document. It loosely corresponds to
a C<< <w:r> >> node in OOXML, but may also contain an anonymous XML
fragment which is the part of the document just before the C<< <w:r> >> 
node -- used for reconstructing the complete document after having changed
the contents of some runs.


=head1 METHODS

=head2 new

  my $run = MsOffice::Word::Surgeon::Run(
    xml_before  => $xml_string,
    props       => $properties_string,
    inner_texts => [MsOffice::Word::Surgeon::Text(...), ...],
  );

Constructor for a new run object. Arguments are :

=over

=item xml_before

A string containing arbitrary XML preceding that run in the complete document.
The string may be empty but must be present.

=item props

A string containing XML for the properties of this run (for example instructions
for bold, italic, font, etc.). The module does not parse this information;
it just compares the string for equality with the next run.


=item inner_texts

An array of L<MsOffice::Word::Surgeon::Text> objects, corresponding to the
XML C<< <w:t> >> nodes inside the run.

=back

=head2 as_xml

  my $xml = $run->as_xml;

Returns the XML representation of that run.


=head2 merge

  $run->merge($next_run);

Merge the contents of C<$next_run> together with the current run.
This is only possible if both runs have the same properties (same
string returned by the C<props> method), and if the next run has
an empty C<xml_before> attribute; if the conditions are not met,
an exception is raised.


