package MsOffice::Word::Surgeon::Text;
use feature 'state';
use Moose;
use MsOffice::Word::Surgeon::Utils qw(maybe_preserve_spaces);
use Carp                           qw(croak);

use namespace::clean -except => 'meta';

has 'xml_before'   => (is => 'ro', isa => 'Str', required => 1);
has 'literal_text' => (is => 'ro', isa => 'Str', required => 1);

sub as_xml {
  my $self = shift;

  my $xml = $self->xml_before;
  if (my $lit_txt = $self->literal_text) {
    my $space_attr  = maybe_preserve_spaces($lit_txt);
    $xml .= "<w:t$space_attr>$lit_txt</w:t>";
  }
  return $xml;
}



sub merge {
  my ($self, $next_text) = @_;

  !$next_text->xml_before
    or croak "cannot merge -- next text contains xml before the text : "
           . $next_text->xml_before;

  $self->{literal_text} .= $next_text->literal_text;

}



sub replace {
  my ($self, $pattern, $replacement_callback, %replacement_args) = @_;

  my @fragments = split qr[($pattern)], $self->{literal_text}, -1;

  my $xml = "";
  my $is_first = 1;

  while (my ($txt_before, $matched) = splice (@fragments, 0, 2)) {
    my $xml_before = $is_first_fragment ? $self->xml_before : "";
    if ($txt_before) {
      my $run = MsOffice::Word::Surgeon::Run->new(
        xml_before  => '',
        props       => $replacement_args{run}->props,
        inner_texts => [MsOffice::Word::Surgeon::Text->new(
                           xml_before   => $xml_before,
                           literal_text => $txt_before,
                          )],
       );
      $xml .= $run->as_xml;

      $xml_before = "";
    }
    if ($matched) {
      $xml .= $replacement_callback->(matched    => $matched,
                                      xml_before => $xml_before,
                                      %replacement_args,
                                     );

    }

    $is_first_fragment = 0;
  }

  return $xml;
}


1;

__END__

=encoding ISO-8859-1

=head1 NAME

MsOffice::Word::Surgeon::Text -- internal representation for a node of literal text

=head1 DESCRIPTION

This is used internally by L<MsOffice::Word::Surgeon> for storing
a chunk of literal text in a MsWord document. It loosely corresponds to
a C<< <w:t> >> node in OOXML, but may also contain an anonymous XML
fragment which is the part of the document just before the C<< <w:t> >> 
node -- used for reconstructing the complete document after having changed
the contents of some text nodes.


=head1 METHODS

=head2 new

  my $text_node = MsOffice::Word::Surgeon::Text(
    xml_before   => $xml_string,
    literal_text => $text_string,
  );

Constructor for a new text object. Arguments are :

=over

=item xml_before

A string containing arbitrary XML preceding that text node in the complete document.
The string may be empty but must be present.


=item literal_text

A string of literal text.

=back



=head2 as_xml

  my $xml = $text_node->as_xml;

Returns the XML representation of that text node.
The attribute C<< xml:space="preserve" >> is automatically added
if the literal text starts of ends with a space character.


=head2 merge

  $text_node->merge($next_text_node);

Merge the contents of C<$next_text_node> together with the current text node.
This is only possible if the next text node has
an empty C<xml_before> attribute; if this condition is not met,
an exception is raised.

=head2 replace

  my $xml = $text_node->replace($pattern, $replacement_callback, %args);

Replaces all occurrences of C<$pattern> within the text node by
a new string computed by C<$replacement_callback>, and returns a new xml
string corresponding to the result of all these replacements. This is the
internal implementation for public method
L<MsOffice::Word::Surgeon/replace>.


