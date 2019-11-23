package MsOffice::Word::Surgeon::Text;
use feature 'state';
use Moose;
use MsOffice::Word::Surgeon::Utils qw(maybe_preserve_spaces is_at_run_level);
use Carp                           qw(croak);

use namespace::clean -except => 'meta';

has 'xml_before'   => (is => 'ro', isa => 'Str');
has 'literal_text' => (is => 'ro', isa => 'Str', required => 1);

sub as_xml {
  my $self = shift;

  my $xml = $self->xml_before // '';
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
  my ($self, $pattern, $replacement, %args) = @_;

  my $xml = "";
  my $text_node;

  my $xml_before = $self->xml_before;

  # closure to make sure that $xml_before is used only once
  my $maybe_xml_before = sub { my @r = $xml_before ? (xml_before => $xml_before) : ();
                               $xml_before = undef;
                               return @r;
                             };

  # closure to create a new text node
  my $mk_new_text = sub {my ($literal_text) = @_;
                         return MsOffice::Word::Surgeon::Text->new(
                           $maybe_xml_before->(),
                           literal_text => $literal_text,
                          );
                         };


  # closure to create a new run node for enclosing a text node
  my $add_new_run = sub { my ($text_node) = @_;
                          my $run = MsOffice::Word::Surgeon::Run->new(
                            xml_before  => '',
                            props       => $args{run}->props,
                            inner_texts => [$text_node],
                           );
                          $xml .= $run->as_xml;
                        };


  my @fragments = split qr[($pattern)], $self->{literal_text}, -1;

  while (my ($txt_before, $matched) = splice (@fragments, 0, 2)) {
    if ($matched) {
      # new text to replace the matched fragment
      my $new_txt = !ref $replacement ? $replacement
                                      :   # invoke the callback sub
                                        $replacement->(matched => $matched,
                                                       (!$txt_before ? $maybe_xml_before->() : ()),
                                                       %args);

      my $new_txt_is_xml = $new_txt =~ /^</;
      if ($new_txt_is_xml) {
        if ($txt_before) {

          # clear up $text_node, if any
          if ($text_node) {
            $xml .= $text_node->as_xml;
            $text_node = undef;
          }

          $add_new_run->($mk_new_text->($txt_before));
        }

        # add the text or xml that replaces the match
        $xml .= $new_txt;

      }
      else { # new_text is literal text
        $text_node //= $mk_new_text->('');
        $text_node->{literal_text} .= ($txt_before // '') . $new_txt;
      }
    }

    else { # just $txt_before, no $matched -- so this is the last loop
      $text_node //= $mk_new_text->('');
      $text_node->{literal_text} .= $txt_before;
    }
  }


  if ($text_node) {
    if (is_at_run_level($xml)) {
      $add_new_run->($text_node);
    }
    else {
      $xml .= $text_node->as_xml;
    }
  }

  return $xml;
}



sub uppercase {
  my $self = shift;
  $self->{literal_text} = uc($self->{literal_text});
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

=head2 uppercase

Puts the literal text within the node into uppercase letters.

