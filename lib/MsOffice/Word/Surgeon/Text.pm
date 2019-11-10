package MsOffice::Word::Surgeon::Text;
use feature 'state';
use Moose;
use MsOffice::Word::Surgeon::Utils qw(maybe_preserve_spaces);

use namespace::clean -except => 'meta';

has 'xml_before'   => (is => 'ro', isa => 'Str', required => 1);
has 'literal_text' => (is => 'ro', isa => 'Str', required => 1);

sub add_literal_text {
  my ($self, $more_text) = @_;
  $self->{literal_text} .= $more_text;
}


sub as_xml {
  my $self = shift;

  my $xml = $self->xml_before;
  if (my $lit_txt = $self->literal_text) {
    my $space_attr  = maybe_preserve_spaces($lit_txt);
    $xml .= "<w:t$space_attr>$lit_txt</w:t>";
  }
  return $xml;
}


1;


