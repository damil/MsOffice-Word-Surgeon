package MsOffice::Word::Surgeon::PackagePart;
use feature 'state';
use Moose;
use MooseX::StrictConstructor;
use MsOffice::Word::Surgeon::Utils qw(maybe_preserve_spaces is_at_run_level);
use XML::LibXML;
use Carp                           qw(croak);
use Encode                         qw(encode_utf8 decode_utf8);

# constant integers to specify indentation modes -- see L<XML::LibXML>
use constant XML_NO_INDENT     => 0;
use constant XML_SIMPLE_INDENT => 1;

use namespace::clean -except => 'meta';

has 'surgeon'       => (is => 'ro', isa => 'MsOffice::Word::Surgeon', required => 1, weak_ref => 1);
has 'part_name'     => (is => 'ro', isa => 'Str',                     required => 1);




our $VERSION = '1.08';



has 'contents'  => (is => 'rw', isa => 'Str',          init_arg => undef,
                    builder => 'original_contents', lazy => 1,
                    trigger => sub {shift->clear_runs});

has 'runs'      => (is => 'ro', isa => 'ArrayRef',     init_arg => undef,
                    builder => '_runs', lazy => 1, clearer => 'clear_runs');




#======================================================================
# GLOBAL VARIABLES
#======================================================================

# Various regexes for removing uninteresting XML information
my %noise_reduction_regexes = (
  proof_checking        => qr(<w:(?:proofErr[^>]+|noProof/)>),
  revision_ids          => qr(\sw:rsid\w+="[^"]+"),
  complex_script_bold   => qr(<w:bCs/>),
  page_breaks           => qr(<w:lastRenderedPageBreak/>),
  language              => qr(<w:lang w:val="[^/>]+/>),
  empty_run_props       => qr(<w:rPr></w:rPr>),
  soft_hyphens          => qr(<w:softHyphen/>),
 );

my @noise_reduction_list = qw/proof_checking revision_ids
                              complex_script_bold page_breaks language 
                              empty_run_props soft_hyphens/;




#======================================================================
# METHODS
#======================================================================

sub zip_member_name {
  my $self = shift;
  return sprintf "word/%s.xml", $self->part_name;
}



sub original_contents { # can also be called later, not only as lazy constructor
  my $self = shift;

  my $bytes      = $self->surgeon->zip->contents($self->zip_member_name)
    or die "no contents for part ", $self->part_name;
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


  # split XML content into run fragments
  my $contents      = $self->contents;
  my @run_fragments = split m[$run_regex], $contents, -1;
  my @runs;

  # build internal RUN objects
 RUN:
  while (my ($xml_before_run, $props, $run_contents) = splice @run_fragments, 0, 3) {
    $run_contents //= '';

    # split XML of this run into text fragmentsn
    my @txt_fragments = split m[$txt_regex], $run_contents, -1;
    my @texts;

    # build internal TEXT objects
  TXT:
    while (my ($xml_before_text, $txt_contents) = splice @txt_fragments, 0, 2) {
      next TXT if !$xml_before_text && ( !(defined $txt_contents) || $txt_contents eq '');
      push @texts, MsOffice::Word::Surgeon::Text->new(
        xml_before   => $xml_before_text // '',
        literal_text => $txt_contents    // '',
       );
    }

    # assemble TEXT objects into a RUN object
    next RUN if !$xml_before_run && !@texts;
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

  # replace break tags by newlines
  $txt =~ s[<w:br/>][\n]g;

  # replace tab nodes by ASCII tabs
  $txt =~ s/<w:tab[^s][^>]*>/\t/g;

  # remove all remaining XML tags
  $txt =~ s/<[^>]+>//g;

  return $txt;
}




#======================================================================
# MODIFYING CONTENTS
#======================================================================

sub cleanup_XML {
  my ($self, @merge_args) = @_;

  $self->reduce_all_noises;
  my @names_of_ASK_fields = $self->unlink_fields;
  $self->suppress_bookmarks(@names_of_ASK_fields);
  $self->merge_runs(@merge_args);
}

sub noise_reduction_regex {
  my ($self, $regex_name) = @_;
  my $regex = $noise_reduction_regexes{$regex_name}
    or croak "->noise_reduction_regex('$regex_name') : unknown regex name";
  return $regex;
}

sub reduce_noise {
  my ($self, @noises) = @_;

  # gather regexes to apply, given either directly as regex refs, or as names of builtin regexes
  my @regexes = map {ref $_ eq 'Regexp' ? $_ : $self->noise_reduction_regex($_)} @noises;

  # get contents, apply all regexes, put back the modified contents.
  my $contents = $self->contents;
  no warnings 'uninitialized'; # for regexes without capture groups, $1 will be undef
  $contents =~ s/$_/$1/g foreach @regexes;
  $self->contents($contents);
}

sub reduce_all_noises {
  my $self = shift;

  $self->reduce_noise(@noise_reduction_list);
}

sub suppress_bookmarks {
  my ($self, @names_to_erase) = @_;

  # regex to find bookmarks markup
  state $bookmark_rx = qr{
     <w:bookmarkStart         # initial tag
       .+? w:id="(\d+)"       # 'id' attribute, bookmark identifier -- capture 1
       .+? w:name="([^"]+)"   # 'name' attribute                    -- capture 2
       .*? />                 # end of this tag
       (.*?)                  # bookmark contents (may be empty)    -- capture 3
     <w:bookmarkEnd           # ending tag
       \s+ w:id="\1"          # same 'id' attribute
       .*? />                 # end of this tag
    }sx;

  # closure to decide what to do with bookmark contents
  my %should_erase_contents = map {($_ => 1)} @names_to_erase;
  my $deal_with_bookmark_text = sub {
    my ($bookmark_name, $bookmark_contents) = @_;
    return $should_erase_contents{$bookmark_name} ? "" : $bookmark_contents;
  };

  # remove bookmarks markup
  my $contents = $self->contents;
  $contents    =~ s{$bookmark_rx}{$deal_with_bookmark_text->($2, $3)}eg;

  # re-inject the modified contents
  $self->contents($contents);
}

sub merge_runs {
  my ($self, %args) = @_;

  # check validity of received args
  state $is_valid_arg = {no_caps => 1};
  $is_valid_arg->{$_} or croak "merge_runs(): invalid arg: $_"
    foreach keys %args;

  my @new_runs;
  # loop over internal "run" objects
  foreach my $run (@{$self->runs}) {

    $run->remove_caps_property if $args{no_caps};

    # check if the current run can be merged with the previous one
    if (   !$run->xml_before                    # no other XML markup between the 2 runs
        && @new_runs                            # there was a previous run
        && $new_runs[-1]->props eq $run->props  # both runs have the same properties
       ) {
      # conditions are OK, so merge this run with the previous one
      $new_runs[-1]->merge($run);
    }
    else {
      # conditions not OK, just push this run without merging
      push @new_runs, $run;
    }
  }

  # reassemble the whole stuff and inject it as new contents
  $self->contents(join "", map {$_->as_xml} @new_runs);
}





sub unlink_fields {
  my $self = shift;

  # must find out what are the ASK fields before erasing the markup
  state $ask_field_rx = qr[<w:instrText[^>]+?>\s+ASK\s+(\w+)];
  my $contents            = $self->contents;
  my @names_of_ASK_fields = $contents =~ /$ask_field_rx/g;


  # regexes to remove field nodes and "field instruction" nodes
  state $field_instruction_txt_rx = qr[<w:instrText.*?</w:instrText>];
  state $field_boundary_rx        = qr[<w:fldChar
                                         (?:  [^>]*?/>                 # ignore all attributes until end of node ..
                                            |                          # .. or
                                              [^>]*?>.*?</w:fldChar>)  # .. ignore node content until closing tag
                                      ]x;   # field boundaries are encoded as  "begin" / "separate" / "end"
  state $simple_field_rx          = qr[</?w:fldSimple[^>]*>];

  $self->reduce_noise($field_instruction_txt_rx, $field_boundary_rx, $simple_field_rx);

  return @names_of_ASK_fields;
}


sub replace {
  my ($self, $pattern, $replacement_callback, %replacement_args) = @_;

  # cleanup the XML structure so that replacements work better
  my $keep_xml_as_is = delete $replacement_args{keep_xml_as_is};
  $self->cleanup_XML unless $keep_xml_as_is;

  # special option to avoid modying contents
  my $dont_overwrite_contents = delete $replacement_args{dont_overwrite_contents};

  # apply replacements and generate new XML
  my $xml = join "", map {
    $_->replace($pattern, $replacement_callback, %replacement_args)
  }  @{$self->runs};

  $self->contents($xml) unless $dont_overwrite_contents;

  return $xml;
}

#======================================================================
# DELEGATION TO SUBCLASSES
#======================================================================

sub change {
  my $self = shift;

  my $change = MsOffice::Word::Surgeon::Change->new(rev_id => $self->{rev_id}++, @_);
  return $change->as_xml;
}




#======================================================================
# BEFORE SAVING
#======================================================================


sub _update_contents_in_zip {
  my $self = shift;

  $self->surgeon->zip->contents($self->zip_member_name, encode_utf8($self->contents));
}






1;

__END__

=encoding ISO-8859-1

=head1 NAME

MsOffice::Word::Surgeon::PackagePart - TODO

=head1 DESCRIPTION

TODO


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


=head2 replace

  my $xml = $run->replace($pattern, $replacement_callback, %replacement_args);

Replaces all occurrences of C<$pattern> within all text nodes by
a new string computed by C<$replacement_callback>, and returns a new xml
string corresponding to the result of all these replacements. This is the
internal implementation for public method
L<MsOffice::Word::Surgeon/replace>.


=head2 remove_caps_property

Searches in the run properties for a C<< <w:caps/> >> property;
if found, removes it, and replaces all inner texts by their
uppercase equivalents.
