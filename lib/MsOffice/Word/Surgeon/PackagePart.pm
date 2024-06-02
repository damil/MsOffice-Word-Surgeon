package MsOffice::Word::Surgeon::PackagePart;
use 5.24.0;
use Moose;
use MooseX::StrictConstructor;
use MsOffice::Word::Surgeon::Carp;
use MsOffice::Word::Surgeon::Utils qw(maybe_preserve_spaces is_at_run_level parse_attrs decode_entities encode_entities);
use MsOffice::Word::Surgeon::Run;
use MsOffice::Word::Surgeon::Text;
use XML::LibXML;
use List::Util                     qw(max);
use Scalar::Util                   qw(dualvar);
use match::simple                  qw(match);

# syntactic sugar for attributes
sub has_inner ($@) {my $attr = shift; has($attr => @_, lazy => 1, builder => "_$attr", init_arg => undef)}


our $VERSION = '2.06';

# constant integers to specify indentation modes -- see L<XML::LibXML>
use constant XML_NO_INDENT     => 0;
use constant XML_SIMPLE_INDENT => 1;

use namespace::clean -except => 'meta';



#======================================================================
# ATTRIBUTES
#======================================================================


# attributes passed to the constructor
has       'surgeon'        => (is => 'ro', isa => 'MsOffice::Word::Surgeon', required => 1, weak_ref => 1);
has       'part_name'      => (is => 'ro', isa => 'Str',                     required => 1);


# attributes constructed by the module -- not received through the constructor
has_inner 'contents'       => (is => 'rw', isa => 'Str',      trigger => \&_on_new_contents);
has_inner 'runs'           => (is => 'ro', isa => 'ArrayRef', clearer => 'clear_runs');
has_inner 'relationships'  => (is => 'ro', isa => 'ArrayRef');
has_inner 'images'         => (is => 'ro', isa => 'HashRef');

has 'contents_has_changed' => (is => 'bare', isa => 'Bool', default => 0);
has 'was_cleaned_up'       => (is => 'bare', isa => 'Bool', default => 0);

#======================================================================
# GLOBAL VARIABLES
#======================================================================

# Various regexes for removing uninteresting XML information
my %noise_reduction_regexes = (
  proof_checking         => qr(<w:(?:proofErr[^>]+|noProof/)>),
  revision_ids           => qr(\sw:rsid\w+="[^"]+"),
  complex_script_bold    => qr(<w:bCs/>),
  page_breaks            => qr(<w:lastRenderedPageBreak/>),
  language               => qr(<w:lang w:val="[^/>]+/>),
  empty_run_props        => qr(<w:rPr></w:rPr>),
  soft_hyphens           => qr(<w:softHyphen/>),
 );

my @noise_reduction_list = qw/proof_checking revision_ids
                              complex_script_bold page_breaks language
                              empty_run_props soft_hyphens/;

#======================================================================
# LAZY ATTRIBUTE CONSTRUCTORS AND TRIGGERS
#======================================================================


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
      next TXT if !$xml_before_text && !$txt_contents;
      $_ //= '' for $xml_before_text, $txt_contents;
      decode_entities($txt_contents);
      push @texts, MsOffice::Word::Surgeon::Text->new(xml_before   => $xml_before_text,
                                                      literal_text => $txt_contents);
    }

    # assemble TEXT objects into a RUN object
    next RUN if !$xml_before_run && !@texts;
    $_ //= '' for $xml_before_run, $props;
    push @runs, MsOffice::Word::Surgeon::Run->new(xml_before  => $xml_before_run,
                                                  props       => $props,
                                                  inner_texts => \@texts);
  }

  return \@runs;
}


sub _relationships {
  my $self = shift;

  # xml that describes the relationships for this package part
  my $rel_xml = $self->_rels_xml;

  # parse the relationships and assemble into a sparse array indexed by relationship ids
  my @relationships;
  while ($rel_xml =~ m[<Relationship\s+(.*?)/>]g) {
    my %attrs = parse_attrs($1);
    $attrs{$_} or croak "missing attribute '$_' in <Relationship> node" for qw/Id Type Target/;
    ($attrs{num}        = $attrs{Id})  =~ s[^\D+][];
    ($attrs{short_type} = $attrs{Type}) =~ s[^.*/][];
    $relationships[$attrs{num}] = \%attrs;
  }

  return \@relationships;
}


sub _images {
  my $self = shift;

  # get relationship ids associated with images
  my %rel_image  = map  {$_->{Id} => $_->{Target}}
                   grep {$_ && $_->{short_type} eq 'image'}
                   $self->relationships->@*;

  # get titles and relationship ids of images found within the part contents
  my %image;
  my @drawings = $self->contents =~ m[<w:drawing>(.*?)</w:drawing>]g;
 DRAWING:
  foreach my $drawing (@drawings) {
    if ($drawing =~ m[<wp:docPr \s+ (.*?) />
                      .*?
                      <a:blip \s+ r:embed="(\w+)"]x) {
      my ($lst_attrs, $rId) = ($1, $2);
      my %attrs = parse_attrs($lst_attrs);
      my $img_id = $attrs{title} || $attrs{descr}
        or next DRAWING;

      $image{$img_id} = "word/$rel_image{$rId}"
        or die "couldn't find image for relationship '$rId' associated with image '$img_id'";
        # NOTE: targets in the rels XML miss the "word/" prefix, I don't know why.
    }
  }

  return \%image;
}


sub _contents {shift->original_contents}

sub _on_new_contents {
  my $self = shift;

  $self->clear_runs;
  $self->{contents_has_changed} = 1;
  $self->{was_cleaned_up}       = 0;
}

#======================================================================
# GENERAL METHODS
#======================================================================


sub  _rels_xml { # rw accessor
  my ($self, $new_xml) = @_;
  my $rels_name = sprintf "word/_rels/%s.xml.rels", $self->part_name;
  return $self->surgeon->xml_member($rels_name, $new_xml);
}


sub zip_member_name {
  my $self = shift;
  return sprintf "word/%s.xml", $self->part_name;
}


sub original_contents {
  my $self = shift;

  return $self->surgeon->xml_member($self->zip_member_name);
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

  # avoid doing it twice
  return if $self->{was_cleaned_up};

  # start the cleanup
  $self->reduce_all_noises;

  # remember the names of ASK fields before erasing them
  my $contents            = $self->contents;
  my @names_of_ASK_fields = $contents =~ m[<w:instrText[^>]+?>\s+ASK\s+(\w+)]g;

  # unlink fields, suppress bookmarks (and suppress content of ASK bookmarks), merge runs
  $self->unlink_fields;
  $self->suppress_bookmarks(full_range => \@names_of_ASK_fields, markup_only => qr/./);
  $self->merge_runs(@merge_args);

  # flag the fact that the cleanup was done
  $self->{was_cleaned_up} = 1;
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

sub noise_reduction_regex {
  my ($self, $regex_name) = @_;
  my $regex = $noise_reduction_regexes{$regex_name}
    or croak "->noise_reduction_regex('$regex_name') : unknown regex name";
  return $regex;
}

sub reduce_all_noises {
  my $self = shift;

  $self->reduce_noise(@noise_reduction_list);
}



sub merge_runs {
  my ($self, %args) = @_;

  # check validity of received args
  state $is_valid_arg = {no_caps => 1};
  my @invalid_args    = grep {!$is_valid_arg->{$_}} keys %args;
  croak "merge_runs(): invalid arg(s): " . join ", ", @invalid_args if @invalid_args;

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



sub replace {
  my ($self, $pattern, $replacement_callback, %replacement_args) = @_;

  # shared initial string for error messages
  my $error_msg = '->replace($pattern, $callback, %args)';

  # default value for arg 'cleanup_XML', possibly from deprecated arg 'keep_xml_as_is'
  if (delete $replacement_args{keep_xml_as_is}) {
    not exists $replacement_args{cleanup_XML}
      or croak "$error_msg: deprecated arg 'keep_xml_as_is' conflicts with arg 'cleanup_XML'";
    carp "$error_msg: arg 'keep_xml_as_is' is deprecated, use 'cleanup_XML' instead";
    $replacement_args{cleanup_XML} = 0;
  }
  else {
    $replacement_args{cleanup_XML} //= 1; # default
  }

  # cleanup the XML structure so that replacements work better
  if (my $cleanup_args = $replacement_args{cleanup_XML}) {
    $cleanup_args = {} if ! ref $cleanup_args;
    ref $cleanup_args eq 'HASH'
      or croak "$error_msg: arg 'cleanup_XML' should be a hashref";
    $self->cleanup_XML(%$cleanup_args);
  }

  # check for presences of a special option to avoid modying contents
  my $dont_overwrite_contents = delete $replacement_args{dont_overwrite_contents};

  # apply replacements and generate new XML
  my $xml = join "",
            map {$_->replace($pattern, $replacement_callback, %replacement_args)} $self->runs->@*;

  # overwrite previous contents
  $self->contents($xml) unless $dont_overwrite_contents;

  return $xml;
}


sub _update_contents_in_zip { # called for each part before saving the zip file
  my $self = shift;

  $self->surgeon->xml_member($self->zip_member_name, $self->contents)
    if $self->{contents_has_changed};
}



#======================================================================
# OPERATIONS ON BOOKMARKS
#======================================================================


sub _split_into_bookmark_nodes {
  my ($self, $xml) = @_;

  # regex to find bookmark tags
  state $bookmark_rx = qr{
     (                               # the whole tag                       -- capture 1
      <w:bookmark(Start|End)         # kind of tag name                    -- capture 2
        .+?                          # optional attributes (may be w:colFirst, w:colLast) -- no capture
        w:id="(\d+)"                 # 'id' attribute, bookmark identifier -- capture 3
        (?: \h+ w:name="([^"]+)")?   # optional 'name' attribute           -- capture 4
        \h* />                       # end of this tag
     )                               # end of capture 1
    }sx;

  # split the whole xml according to the regex. Captured groups are also added to the stack
  my @xml_chunks = split /$bookmark_rx/, $xml;

  # walk through the list of fragments and build a stack of hashrefs as bookmark nodes
  my @bookmark_nodes;
  while (my @chunk = splice @xml_chunks, 0, 5) {
    my %node;  @node{qw/xml_before node_xml node_kind id name/} = @chunk; # initialize a node hash
    $node{$_} //= "" for qw/xml_before node_kind node_xml/;               # empty strings instead of undef
    push @bookmark_nodes, \%node;
  }
  # note : in most cases the last "node" is not really a node : it has no 'node_kind', but only 'xml_before'
  
  # return the stack
  return @bookmark_nodes;
}



sub suppress_bookmarks {
  my ($self, %options) = @_;

  # check if options are valid and supply defaults
  my @invalid_opt = grep {!/^(full_range|markup_only)$/} keys %options;
  croak "suppress_bookmarks: invalid options: " . join(", ", @invalid_opt) if @invalid_opt;
  %options = (markup_only => qr/./) if ! keys %options;

  
  # loop on bookmark nodes
  my @bookmark_nodes = $self->_split_into_bookmark_nodes($self->contents);
  my %node_ix_by_id;
  while (my ($ix, $node) = each @bookmark_nodes) {
    if ($node->{node_kind} eq 'Start') {
      $node_ix_by_id{$node->{id}} = $ix;
    }
    elsif ($node->{node_kind} eq 'End') {

      # try to find the corresponding bookmarkStart node. 
      my $start_ix = $node_ix_by_id{$node->{id}};

      if (!defined $start_ix) {
        # no start node -- this is because it was within a field that has been erased. So just clear the end node.
        $node->{node_xml} = "";
      }

      else {
        my $start_node         = $bookmark_nodes[$start_ix];
        my $bookmark_name      = $start_node->{name};
        my $should_erase_range = match($bookmark_name, $options{full_range});
        if ($should_erase_range || match($bookmark_name, $options{markup_only})) {

          # erase markup (start and end nodes)
          $node->{node_xml}       = "";
          $start_node->{node_xml} = "";

          # if required, erase inner range
          if ($should_erase_range) {
            for my $erase_ix ($start_ix+1 .. $ix) {
              my $inner_node = $bookmark_nodes[$erase_ix];
              !$inner_node->{node_xml}
                or die "cannot erase contents of bookmark '$bookmark_name' "
                . "because it contains the start of bookmark '$inner_node->{name}'";
              $inner_node->{xml_before} = "";
            }
          }
        }
      }
    }
  }

  # re-build the whole XML from all remaining fragments, and inject it back
  my $new_contents = join "", map {@{$_}{qw/xml_before node_xml/}} @bookmark_nodes;
  $self->contents($new_contents);
}


 sub reveal_bookmarks {
  my ($self, @marking_args) = @_;

  # auxiliary objects
  my $marker            = MsOffice::Word::Surgeon::PackagePart::_BookmarkMarker->new(@marking_args);
  my $paragraph_tracker = MsOffice::Word::Surgeon::PackagePart::_ParaTracker->new;

  # loop over bookmark nodes
  my @bookmark_name_by_id;

  my @bookmark_nodes    = $self->_split_into_bookmark_nodes($self->contents);
  foreach my $node (@bookmark_nodes) {

    # count opening and closing paragraphs in xml before this node
    $paragraph_tracker->count_paragraphs($node->{xml_before});

    # add visible runs before or after bookmark nodes
    if ($node->{node_kind} eq 'Start') {
      $bookmark_name_by_id[$node->{id}] = $node->{name};
      substr $node->{node_xml}, 0, 0, $paragraph_tracker->maybe_add_paragraph($marker->mark($node->{name}, 0));
    }
    elsif ($node->{node_kind} eq 'End') {
      my $bookmark_name  = $bookmark_name_by_id[$node->{id}];
      $node->{node_xml} .= $paragraph_tracker->maybe_add_paragraph($marker->mark($bookmark_name, 1));
    }
  }

  # re-build the whole XML and inject it back
  my $new_contents = join "", map {@{$_}{qw/xml_before node_xml/}} @bookmark_nodes;
  $self->contents($new_contents);
}


#======================================================================
# OPERATIONS ON FIELDS
#======================================================================



sub _split_into_field_nodes {
  my ($self, $xml) = @_;

  # regex to find field nodes
  state $field_rx = qr{
    < w:fld                   # initial prefix for a field node
      (Simple|Char)           # $1 : distinguish between simple fields and complex fields
      \h* ([^>]*?)            # $2 : node attributes
      (?:                     # either ..
          />                  # .. the end of an empty XML element
        |                     # or ..
          >                   # .. the end of the opening tag
            (.*?)             # .. $3: some node content
          </w:fld\g1>         # .. the closing tag
      )
    }sx;

  # split the whole xml according to the regex. Captured groups are also added to the stack
  my @xml_chunks = split /$field_rx/, $xml;

  # walk through the list of fragments and build a stack of hashrefs as field nodes
  my @field_stack;

 NODE:
  while (my @chunk = splice @xml_chunks, 0, 4) {

    # initialize a node hash
    my %node;  @node{qw/xml_before field_kind attrs node_content/} = @chunk;
    $node{$_} //= "" for qw/xml_before field_kind attrs node_content/;

    # node attributes
    my %attrs = parse_attrs($node{attrs});

    if ($node{field_kind} eq 'Simple') {
      # for a simple field, all information is within the XML node
      push @field_stack, {xml_before => $node{xml_before},
                          code       => $attrs{instr},
                          result     => $node{node_content},
                          status     => "end"};
    }

    elsif ($node{field_kind} eq 'Char') {
      # for a complex field, we need an auxiliary subroutine to handle the begin/separate/end parts
      $self->_handle_fldChar_node(\@field_stack, \%node, \%attrs);
    }
        
    else {
      # this is the last node, it only contains 'xml_before'
      push @field_stack, {xml_before => $node{xml_before},
                          code       => '',
                          result     => '',
                          status     => "end"};
    }

  }
  
  # return the stack
  return @field_stack;
}



sub _handle_fldChar_node {
  my ($self, $field_stack, $node, $attrs) = @_;

  my $fldChar_type = $attrs->{"w:fldCharType"};

  # a) beginning a of a field : push a new field node on top of the stack
  if ($fldChar_type eq 'begin') {
    push @$field_stack,  {xml_before => $node->{xml_before},
                          code       => '',
                          result     => '',
                          status     => "begin"};
  }
    
  # b) frontier between the "code" part and the "result" part : assemble code instructions seen so far
  elsif ($fldChar_type eq 'separate') {
    my $current_field = $field_stack->[-1];
    $current_field && $current_field->{status} eq "begin"
      or croak q{met <w:fldChar w:fldCharType="separate"> without an open w:fldCharType="begin"};
    $current_field->{status} = "separate";
    my @instr_text = $node->{xml_before} =~ m{<w:instrText.*?>(.*?)</w:instrText>}g;
    $current_field->{code} .= join "", @instr_text;
  }

  # c) end of the field : assemble the "result" part.
  # If this field is embedded within the "code" or "result" part of a parent field, pop from stack
  # and insert current content into the parent.
  elsif ($fldChar_type eq 'end') {
    my $current_field = $field_stack->[-1];
    $current_field && $current_field->{status} ne "end"
      or croak q{met <w:fldChar w:fldCharType="end"> without an open w:fldCharType="begin"};
    $current_field->{status}  = "end";
    $current_field->{result} .= $node->{xml_before};

    # lookup previous field on stack; its status tells us if the current field is embedded or not
    if (my $prev_field = $field_stack->[-2]) {
      if ($prev_field->{status} eq 'begin') {
        # the current field is embedded within the "code" part of a parent field
        $prev_field->{code} .= sprintf $self->surgeon->show_embedded_field, $current_field->{code};
        pop @$field_stack;
      }

      elsif ($prev_field->{status} eq 'separate') {
        # the current field is embedded within the "result" part of a parent field
        $prev_field->{result} .= $current_field->{xml_before} . $current_field->{result};
        pop @$field_stack;
      }

      # elsif ($prev_field->{status} eq 'end') : this is an independend field, just leave it on top of stack

    }
  }
}


sub replace_fields {
  my ($self, $field_replacer) = @_;

  my @fields = $self->_split_into_field_nodes($self->contents);

  my @xml_parts = map {($_->{xml_before}, 
                        $_->{code} ? $field_replacer->($_->{code}, $_->{result}) : "")} @fields;

  $self->contents(join "", @xml_parts);
}


sub reveal_fields {
  my $self = shift;

  # replace all fields by a textual representatio of their "code" part
  my $revealer = sub {my ($code, $result) = @_; encode_entities($code); "<w:t>{$code}</w:t>"};
  $self->replace_fields($revealer);
}



sub unlink_fields {
  my $self = shift;

  # replace all fields by just their "result" part (in other words, ignore the "code" part)
  my $unlinker = sub {my ($code, $result) = @_; $result};
  $self->replace_fields($unlinker);
}



#======================================================================
# OPERATIONS ON IMAGES
#======================================================================


sub replace_image {
  my ($self, $image_title, $image_PNG_content) = @_;

  my $member_name = $self->images->{$image_title}
    or die "could not find an image with title: $image_title";
  $self->surgeon->zip->contents($member_name, $image_PNG_content);
}



sub add_image {
  my ($self, $image_PNG_content) = @_;

  # compute a fresh image number and a fresh relationship id
  my @image_members = $self->surgeon->zip->membersMatching(qr[^word/media/image]);
  my @image_nums    = map {$_->fileName =~ /(\d+)/} @image_members;
  my $last_img_num  = max @image_nums // 0;
  my $target        = sprintf "media/image%d.png", $last_img_num + 1;
  my $last_rId_num  = $self->relationships->$#*;
  my $rId           = sprintf "rId%d", $last_rId_num + 1;

  # assemble XML for the new relationship
  my $type          = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
  my $new_rel_xml   = qq{<Relationship Id="$rId" Type="$type" Target="$target"/>};

  # update the rels member
  my $xml = $self->_rels_xml;
  $xml =~ s[</Relationships>][$new_rel_xml</Relationships>];
  $self->_rels_xml($xml);

  # add the image as a new member into the archive
  my $member_name = "word/$target";
  $self->surgeon->zip->addString(\$image_PNG_content, $member_name);

  # update the global content_types if it doesn't include PNG
  my $ct = $self->surgeon->_content_types;
  if ($ct !~ /Extension="png"/) {
    $ct =~ s[(<Types[^>]+>)][$1<Default Extension="png" ContentType="image/png"/>];
    $self->surgeon->_content_types($ct);
  }

  # return the relationship id
  return $rId;
}




#======================================================================
# INTERNAL CLASS FOR TRACKING PARAGRAPHS
#======================================================================

package # hide from PAUSE
        MsOffice::Word::Surgeon::PackagePart::_ParaTracker;
use strict;
use warnings;

sub new {
  my $class = shift;
  my $nb_para = 0;
  bless \$nb_para, $class;
}

sub count_paragraphs {
  my ($self, $xml) = @_;

  # count opening and closing paragraph nodes
  while ($xml =~  m[<(/)?w:p.*?(/)?>]g) {
    next if $2; # self-ending node -- doesn't change the number of paragraphs
    $$self += $1 ? -1 : +1;
  }
}

sub maybe_add_paragraph {
  my ($self, $xml) = @_;  

  # add paragraph nodes only if the ParaTracker is currently outside of any paragraph
  my $is_outside_para = !$$self;
  return $is_outside_para && $xml ? "<w:p>$xml</w:p>" : $xml;
};
  

#======================================================================
# INTERNAL CLASS FOR INTRODUCING BOOKMARK MARKERS
#======================================================================

package # hide from PAUSE
        MsOffice::Word::Surgeon::PackagePart::_BookmarkMarker;
use strict;
use warnings;
use MsOffice::Word::Surgeon::Utils qw(encode_entities);
use Carp                           qw(croak carp);

sub new {
  my $class = shift;
  my %self = @_;

  $self{color} //= "yellow";
  $self{props} //=  qq{<w:highlight w:val="%s"/>};
  $self{start} //=  "<%s>";
  $self{end}   //=  "</%s>";
  $self{ignore}  = qr/^_/ if not exists $self{ignore};

  $self{color} =~ m{ ^( black     | blue | cyan | darkBlue    | darkCyan |
                        darkGray  | darkGreen   | darkMagenta | darkRed  | darkYellow | green |
                        lightGray | magenta     | none        | red      | white      | yellow )$}x
    or carp "invalid color : $self{color}";                          
                          
  bless \%self, $class;
}

sub mark {
  my ($self, $bookmark_name, $is_end_node) = @_;

  # some bookmarks are just ignored
  return ""
    if $self->{ignore} and $bookmark_name =~ $self->{ignore};

  # build the visible text
  no warnings 'redundant'; # because sprintf templates may decide not to use their arguments
  my $sprintf_node = $is_end_node ? $self->{end} : $self->{start};
  my $text         = sprintf $sprintf_node, $bookmark_name;
  my $props        = sprintf $self->{props}, $self->{color};
  encode_entities($text);

  # full xml for a visible run before or after the boookmark node
  return "<w:r><w:rPr>$props</w:rPr><w:t>$text</w:t></w:r>";
}


1;

__END__

=encoding ISO-8859-1

=head1 NAME

MsOffice::Word::Surgeon::PackagePart - Operations on a single part within the ZIP package of a docx document

=head1 SYNOPSIS

  my $part = $surgeon->document;
  print $part->plain_text;
  $part->replace(qr[$pattern], $replacement_callback);
  $part->replace_image($image_alt_text, $image_PNG_content);
  $part->unlink_fields;
  $part->reveal_bookmarks;


=head1 DESCRIPTION

This class is part of L<MsOffice::Word::Surgeon>; it encapsulates operations for a single
I<package part> within the ZIP package of a C<.docx> document.
It is mostly used for the I<document> part, that contains the XML representation of the
main document body. However, other parts such as headers, footers, footnotes, etc. have the
same internal representation and therefore the same operations can be invoked.


=head1 METHODS

=head2 new

  my $part = MsOffice::Word::Surgeon::PackagePart->new(
    surgeon   => $surgeon,
    part_name => $name,
  );

Constructor for a new part object. This is called internally from
L<MsOffice::Word::Surgeon>; it is not meant to be called directly
by clients. 

=head3 Constructor arguments


=over

=item surgeon

a weak reference to the main surgeon object 

=item part_name

ZIP member name of this part

=back

=head3 Other attributes

Other attributes, not passed through the constructor but generated lazily on demand, are :

=over

=item contents

the XML contents of this part

=item runs

a decomposition of the XML contents into a collection of
L<MsOffice::Word::Surgeon::Run> objects.

=item relationships

an arrayref of Office relationships associated with this part. This information comes from
a C<.rels> member in the ZIP archive, named after the name of the package part.
Array indices correspond to relationship numbers. Array values are hashrefs with
keys

=over

=item Id

the full relationship id

=item num

the numeric part of C<rId>

=item Type

the full reference to the XML schema for this relationship

=item short_type

only the last word of the type, e.g. 'image', 'style', etc.

=item Target

designation of the target within the ZIP file. The prefix 'word/' must be
added for having a complete Zip member name.

=back



=item images

a hashref of images within this package part. Keys of the hash are image I<alternative texts>.
If present, the alternative I<title> will be prefered; otherwise the alternative I<description> will be taken
(note : the I<title> field was displayed in Office 2013 and 2016, but more recent versions only display
the I<description> field -- see
L<https://support.microsoft.com/en-us/office/add-alternative-text-to-a-shape-picture-chart-smartart-graphic-or-other-object-44989b2a-903c-4d9a-b742-6a75b451c669|MsOffice documentation>).

Images without alternative text will not be accessible through the current Perl module.

Values of the hash are zip member names for the corresponding
image representations in C<.png> format.


=back


=head2 Contents restitution

=head3 contents

Returns a Perl string with the current internal XML representation of the part
contents.

=head3 original_contents

Returns a Perl string with the XML representation of the
part contents, as it was in the ZIP archive before any
modification.

=head3 indented_contents

Returns an indented version of the XML contents, suitable for inspection in a text editor.
This is produced by L<XML::LibXML::Document/toString> and therefore is returned as an encoded
byte string, not a Perl string.

=head3 plain_text

Returns the text contents of the part, without any markup.
Paragraphs and breaks are converted to newlines, all other formatting instructions are ignored.


=head3 runs

Returns a list of L<MsOffice::Word::Surgeon::Run> objects. Each of
these objects holds an XML fragment; joining all fragments
restores the complete document.

  my $contents = join "", map {$_->as_xml} $self->runs;


=head2 Modifying contents


=head3 cleanup_XML

  $part->cleanup_XML(%args);

Apply several other methods for removing unnecessary nodes within the internal
XML. This method successively calls L</reduce_all_noises>, L</unlink_fields>,
L</suppress_bookmarks> and L</merge_runs>.

Currently there is only one legal arg :

=over

=item C<no_caps>

If true, the method L<MsOffice::Word::Surgeon::Run/remove_caps_property> is automatically
called for each run object. As a result, all texts within runs with the C<caps> property are automatically
converted to uppercase.

=back



=head3 reduce_noise

  $part->reduce_noise($regex1, $regex2, ...);

This method is used for removing unnecessary information in the XML
markup.  It applies the given list of regexes to the whole document,
suppressing matches.  The final result is put back into 
C<< $self->contents >>. Regexes may be given either as C<< qr/.../ >>
references, or as names of builtin regexes (described below).  Regexes
are applied to the whole XML contents, not only to run nodes.


=head3 noise_reduction_regex

  my $regex = $part->noise_reduction_regex($regex_name);

Returns the builtin regex corresponding to the given name.
Known regexes are :

  proof_checking       => qr(<w:(?:proofErr[^>]+|noProof/)>),
  revision_ids         => qr(\sw:rsid\w+="[^"]+"),
  complex_script_bold  => qr(<w:bCs/>),
  page_breaks          => qr(<w:lastRenderedPageBreak/>),
  language             => qr(<w:lang w:val="[^/>]+/>),
  empty_run_props      => qr(<w:rPr></w:rPr>),
  soft_hyphens         => qr(<w:softHyphen/>),

=head3 reduce_all_noises

  $part->reduce_all_noises;

Applies all regexes from the previous method.


=head3 merge_runs

  $part->merge_runs(no_caps => 1); # optional arg

Walks through all runs of text within the document, trying to merge
adjacent runs when possible (i.e. when both runs have the same
properties, and there is no other XML node inbetween).

This operation is a prerequisite before performing replace operations, because
documents edited in MsWord often have run boundaries across sentences or
even in the middle of words; so regex searches can only be successful if those
artificial boundaries have been removed.

If the argument C<< no_caps => 1 >> is present, the merge operation
will also convert runs with the C<w:caps> property, putting all letters
into uppercase and removing the property; this makes more merges possible.



=head3 replace

  $part->replace($pattern, $replacement, %replacement_args);

Replaces all occurrences of C<$pattern> regex within the text nodes by the
given C<$replacement>. This is not exactly like a search-replace
operation performed within MsWord, because the search does not cross boundaries
of text nodes. In order to maximize the chances of successful replacements,
the L</cleanup_XML> method is automatically called before starting the operation.

The argument C<$pattern> can be either a string or a reference to a regular expression.
It should not contain any capturing parentheses, because that would perturb text
splitting operations.

The argument C<$replacement> can be either a fixed string, or a reference to
a callback subroutine that will be called for each match.


The C<< %replacement_args >> hash can be used to pass information to the callback
subroutine. That hash will be enriched with three entries :

=over

=item matched

The string that has been matched by C<$pattern>.

=item run

The run object in which this text resides.

=item xml_before

The XML fragment (possibly empty) found before the matched text .

=back

The callback subroutine may return either plain text or structured XML.
See L<MsOffice::Word::Surgeon::Run/SYNOPSIS> for an example of a replacement callback.

The following special keys within C<< %replacement_args >> are interpreted by the 
C<replace()> method itself, and therefore are not passed to the callback subroutine :

=over

=item keep_xml_as_is

if true, no call is made to the L</cleanup_XML> method before performing the replacements

=item dont_overwrite_contents

if true, the internal XML contents is not modified in place; the new XML after performing
replacements is merely returned to the caller.

=item cleanup_args

the argument should be an arrayref and will be passed to the L</cleanup_XML> method. This
is typically used as 

  $part->replace($pattern, $replacement, cleanup_args => [no_caps => 1]);

=back


=head2 Operations on bookmarks

=head3 suppress_bookmarks

  $part->suppress_bookmarks(full_range => [qw/foo bar/], markup_only => qr/^_/);

Suppresses bookmarks according to the specified options :

=over

=item full_range

For bookmark names matching this option, the bookmark will be fully
suppressed (not only the start and end markers, but also any
content inbetween).


=item markup_only

For bookmark names matching this option, start and end markers
are suppressed, but the inner content remains.

=back

Options may be specified as lists of strings, or regexes, or coderefs ... anything suitable
to be compared through L<match::simple>. In absence of any options, the default
is C<< markup_only => qr/./ >>, meaning that all bookmarks markup is suppressed.

Removing bookmarks is useful because
MsWord may silently insert bookmarks in unexpected places; therefore
some searches within the text may fail because of such bookmarks.

The C<full_range> option is especially convenient for removing bookmarks associated
with ASK fields. Such bookmarks contain ranges of text that are 
never displayed by MsWord.


=head3 reveal_bookmarks

  $part->reveal_bookmarks(color => 'green');

Usually bookmarks boundaries in MsWord are not visible; the only way to have a visual clue is to turn on
an option in
L<https://support.microsoft.com/en-gb/office/troubleshoot-bookmarks-9cad566f-913d-49c6-8d37-c21e0e8d6db0|Advanced / Show document content / Show bookmarks> -- but this only displays where bookmarks start and end, without the names of the bookmarks.

The C<reveal_bookmarks()> method will insert a visible run before each bookmark start and after each bookmark end, showing
the bookmark name. This is an interesting tool for documenting where bookmarks are located in an existing document.

Options to this method are :

=over

=item color

The highligting color for visible marks. This should be a valid
highlighting color, i.e black, blue, cyan, darkBlue, darkCyan,
darkGray, darkGreen, darkMagenta, darkRed, darkYellow, green,
lightGray, magenta, none, red, white or yellow. Default is yellow.

=item props

A string in C<sprintf> format for building the XML to be inserted in C<< <w:rPr> >> node
when displaying bookmarks marks, i.e. the style for displaying such marks.
The default is just a highlighting property :  C<< <w:highlight w:val="%s"/> >>.

=item start

A string in C<sprintf> format for generating text before a bookmark start.
Default is C<< <%s> >>.

=item end

A string in C<sprintf> format for generating text after a bookmark end.
Default is C<< </%s> >>.

=item ignore

A regexp for deciding wich bookmarks will not be revealed. Default is C<< qr/^_/ >>,
because bookmarks with an initial underscore are usually technical bookmarks inserted
automatically by MsWord, such as C<_GoBack> or C<_Toc53196147>.


=back


=head2 Operations on fields


=head3 replace_fields

  my $field_replacer = sub {my ($code, $result) = @_; return "...";};
  $part->replace_fields($field_replacer);

Replaces MsWord fields by the product of the C<< $field_replacer >> callback.
The callback receives two arguments :

=over

=item C<$code>

A plain string containing the field's full code instruction, i.e a keyword followed by optional arguments and switches,
including initial and final spaces. Embedded fields are represented in curly braces, like for example

C<< IF { DOCPROPERTY foo } = "bar" "is bar" "is not bar" >>.

=item C<$result>

An XML fragment containing the current value for the field.

=back

The callback should return an XML fragment suitable to be inserted within an MsWord I<run>. 


=head3 reveal_fields

  $part->reveal_fields;

Replaces each field with a textual representation of its code instruction, embedded in curly braces.


=head3 unlink_fields

  $part->unlink_fields;

Replaces each field with its current result, i.e removing the code instruction.
This is the equivalent of performing Ctrl-Shift-F9 in MsWord on the whole document.



=head2 Operations on images


=head3 replace_image

  $part->replace_image($image_alt_text, $image_PNG_content);

Replaces an existing PNG image by a new image. All features of the old image will
be preserved (size, positioning, border, etc.) -- only the image itself will be
replaced. The C<$image_alt_text> must correspond to the I<alternative text> set in Word
for this image.

This operation replaces a ZIP member within the C<.docx> file. If several XML
nodes refer to the I<same> ZIP member, i.e. if the same image is displayed at several
locations, the new image will appear at all locations, even if they do not have the
same alternative text -- unfortunately this module currently has no facility for
duplicating an existing image into separate instances. So if your intent is to only replace
one instance of the image, your original document should contain several distinct copies
of the C<.PNG> file.


=head3 add_image

  my $rId = $part->add_image($image_PNG_content);

Stores the given PNG image within the ZIP file, adds it as a relationship to the
current part, and returns the relationship id. This operation is not sufficient
to  make the image visible in Word : it just stores the image, but you still
have to insert a proper C<drawing> node in the contents XML, using the C<$rId>.
Future versions of this module may offer helper methods for that purpose;
currently it must be done by hand.


=head1 AUTHOR

Laurent Dami, E<lt>dami AT cpan DOT org<gt>

=head1 COPYRIGHT AND LICENSE

Copyright 2019-2024 by Laurent Dami.

This library is free software; you can redistribute it and/or modify
it under the same terms as Perl itself.



