package MsOffice::Word::Surgeon::Run;
use feature 'state';
use Moose;
use Carp qw(croak);

use namespace::clean -except => 'meta';

has 'xml_before'  => (is => 'ro', isa => 'Str'                                    , required => 1);
has 'props'       => (is => 'ro', isa => 'Str'                                    , required => 1);
has 'inner_texts' => (is => 'ro', isa => 'ArrayRef[MsOffice::Word::Surgeon::Text]', required => 1);


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

  $next_run->isa(__PACKAGE__)
    or croak "argument to merge() should be a " . __PACKAGE__;

  $self->props eq $next_run->props
    or croak sprintf "runs have different properties: '%s' <> '%s'",
                      $self->props, $next_run->props;

  !$next_run->xml_before
    or croak "cannot merge -- next run contains xml before the run : " . $next_run->xml_before;

  # NOTE : sanity checks above are redundant with checks performed in Surgeon::merge_runs ..



  foreach my $txt (@{$next_run->inner_texts}) {
    if (@{$self->{inner_texts}} && !$txt->xml_before) {
      $self->{inner_texts}[-1]->add_literal_text($txt->literal_text);
    }
    else {
      push @{$self->{inner_texts}}, $txt;
    }
  }
}


sub replace {
  my ($self, $pattern, $replacement, @context) = @_;

  $_->replace($pattern, $replacement, run => $self, @context)
    foreach @{$self->inner_texts};
}



1;



