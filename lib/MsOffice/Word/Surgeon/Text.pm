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


sub replace {
  my ($self, $pattern, $replacement, %args) = @_;

  my @fragments = split qr[($pattern)], $self->{literal_text}, -1;

  my $xml = "";
  my $is_first = 1;

  while (my ($txt_before, $matched) = splice (@fragments, 0, 2)) {
    my $xml_before_txt = $is_first ? $self->xml_before      : "";
    if ($txt_before) {
      my $run = MsOffice::Word::Surgeon::Run->new(
        xml_before => '',
        props      => $args{run}->props,
        inner_texts => [MsOffice::Word::Surgeon::Text->new(
                           xml_before   => $xml_before_txt,
                           literal_text => $txt_before
                          )],
       );
      $xml .= $run->as_xml;

      $xml_before_txt = "";
    }
    if ($matched) {
      $xml .= $replacement->(matched => $matched,
                             xml_before_txt => $xml_before_txt,
                             %args,
                            );

    }

    $is_first = 0;
  }

  return $xml;
}





1;

__END__







