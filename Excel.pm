package XML::SAXDriver::Excel;

use strict;
use vars qw($VERSION);

$VERSION = '0.01';

use Spreadsheet::ParseExcel;

sub new {
  my ($class, %params) = @_;
  
  %params = (%params, 
            #### Now declare some private vars, we need them blessed with the object, for threaded/multi-process use.
            _row => [], 
            _row_num => -1,
            _last_row_num => 0,
            _last_col => 0
            );  
           
  return bless \%params, $class;
}

sub parse {
  my $self = shift;
  
  ### Reset vars before parsing
  $self->{'_row'} = [];  ## Used to push row values per row
	$self->{'_row_num'} = -1;  ## Set at -1 since rows are counted from 0
	$self->{'_last_row_num'} = 0;  ## Used to save the last row value received
	$self->{'_last_col'} = 0;
		    
  my $args;
  if (@_ == 1 && !ref($_[0])) 
  {
      $args = { Source => { String => shift }};
  }
  else 
  {
      $args = (@_ == 1) ? shift : { @_ };
  }
  
  my $parse_options = { %$self, %$args };
  $self->{ParseOptions} = $parse_options;
  
  if (!defined($parse_options->{Source})
          || !(
          defined($parse_options->{Source}{String})
          || defined($parse_options->{Source}{ByteStream})
          || defined($parse_options->{Source}{SystemId})
          )) 
  {
    die "XML::SAXDriver::CSV: no source defined for parse\n";
  }
  
  if (defined($parse_options->{Handler})) {
      $parse_options->{DocumentHandler} ||= $parse_options->{Handler};
      $parse_options->{DTDHandler} ||= $parse_options->{DTDHandler};
  }
  
  $parse_options->{NewLine} = "\n" unless defined($parse_options->{NewLine});
  $parse_options->{IndentChar} = "\t" unless defined($parse_options->{IndentChar});
      
  $parse_options->{Parser} ||= Spreadsheet::ParseExcel->new(CellHandler => \&cb_routine, Object => $self, NotSetCell => 1);
  
  my ($ioref, @strings);
  if (defined($parse_options->{Source}{SystemId}) 
      || defined($parse_options->{Source}{ByteStream}) ) {
      $ioref = $parse_options->{Source}{ByteStream};
    if (!$ioref) 
    {
      require IO::File;
      $ioref = IO::File->new($parse_options->{Source}{SystemId})
        || die "Cannot open SystemId '$parse_options->{Source}{SystemId}' : $!";
    }
    else
    {
      die ("Cannot use ByteStream to parse a binary Excel file.  You can only use a file by setting SystemId");
    }
              
  }
  elsif (defined $parse_options->{Source}{String}) 
  {
    die ("Cannot use String to parse a binary Excel file.  You can only use a file by setting SystemId");    
  }
  
  my $document = {};
  $parse_options->{Handler}->start_document($document);
  $parse_options->{Handler}->characters({Data => $parse_options->{NewLine}});
  
  my $doc_element = {
              Name => $parse_options->{File_Tag} || "records",
              Attributes => {},
          };

  $parse_options->{Handler}->start_element($doc_element);
  $parse_options->{Handler}->characters({Data => $parse_options->{NewLine}});
  
  ## Parse file or string
  $parse_options->{Parser}->Parse($parse_options->{Source}{SystemId} || $parse_options->{Source}{String});
  
  
  
  
  _print_xml_finish($self);
  
  ### Reset vars after parsing
  $self->{'_row'} = [];  ## Used to push row values per row
	$self->{'_row_num'} = -1;  ## Set at -1 since rows are counted from 0
	$self->{'_last_row_num'} = 0;  ## Used to save the last row value received
  
  $parse_options->{Handler}->end_element($doc_element);
  
  return $parse_options->{Handler}->end_document($document);
  
}

sub cb_routine($$$$$$)
{    
  my ($self, $oBook, $iSheet, $iRow, $iCol, $oCell) = @_;
  
  my $oWkS = $oBook->{Worksheet}[$iSheet];
         
  $self->{ParseOptions}->{Col_Headings} ||= [];

if ($iCol < $oWkS->{MaxCol})
  {
    
    if ($self->{'_last_col'} > $iCol)
  	{
  	  while ($self->{'_last_col'} < $oWkS->{MaxCol})
  	  {
  	    push(@{$self->{'_row'}}, undef);
  	    $self->{'_last_col'}++;    	    
  	  }  	
  	  _print_xml(@_);  	
  	}
    
    if ($self->{'_last_col'} < $iCol)
  	{
  	  while ($self->{'_last_col'} < $iCol)
  	  {
  	    push(@{$self->{'_row'}}, undef);
  	    $self->{'_last_col'}++;    	    
  	  }    	  
  	}
  	
  	  push(@{$self->{'_row'}}, $oCell->Value());
  	  $self->{'_last_row_num'} = $iRow;
  	  $self->{'_last_col'}++;
  	  return;
  	
  	    	
  }

  push(@{$self->{'_row'}}, $oCell->Value());# if $flag == 0;
    
  _print_xml(@_);
  return;
        
}


sub _print_xml
{
  my ($self, $oBook, $iSheet, $iRow, $iCol, $oCell) = @_;  ### Remember self is passed through the Spreadsheet::ParseExcel object
  
  my $oWkS = $oBook->{Worksheet}[$iSheet];
  
  $self->{'_last_row_num'} = $iRow;
      
  
  $self->{'_last_col'} = 0;      
  my $temp_row = $oCell->Value();
  $self->{'_row_num'} = $self->{'_last_row_num'};       
      
              
      if (!@{$self->{ParseOptions}->{Col_Headings}} && !$self->{ParseOptions}->{Dynamic_Col_Headings}) 
      {
              my $i = 1;
              @{$self->{ParseOptions}->{Col_Headings}} = map { "column" . $i++ } @{$self->{'_row'}};                
      }
      elsif (!@{$self->{ParseOptions}->{Col_Headings}} && $self->{ParseOptions}->{Dynamic_Col_Headings})
      {
              @{$self->{ParseOptions}->{Col_Headings}} = @{$self->{'_row'}};
              $self->{'_row'} = [];  ### Clear the @$row array
              return;  ### So that it does not print the column headings as the content of the first node.                
      }
      
      
      my $el = {
        Name => $self->{ParseOptions}->{Parent_Tag} || "record",
        Attributes => {},
      };
      
      $self->{ParseOptions}->{Handler}->characters(
              {Data => $self->{ParseOptions}->{IndentChar} || "\t"
              }
      );
      $self->{ParseOptions}->{Handler}->start_element($el);
      $self->{ParseOptions}->{Handler}->characters({Data => $self->{ParseOptions}->{NewLine}});

      for (my $i = 0; $i <= $#{$self->{ParseOptions}->{Col_Headings}}; $i++) {
          my $column = { Name => $self->{ParseOptions}->{Col_Headings}->[$i], Attributes => {} };
          $self->{ParseOptions}->{Handler}->characters(
                  {Data => $self->{ParseOptions}->{IndentChar} x 2} 
          );
          $self->{ParseOptions}->{Handler}->start_element($column);
          $self->{ParseOptions}->{Handler}->characters({Data => $self->{'_row'}->[$i]});
          $self->{ParseOptions}->{Handler}->end_element($column);
          $self->{ParseOptions}->{Handler}->characters({Data => $self->{ParseOptions}->{NewLine}});
      }

      $self->{ParseOptions}->{Handler}->characters(
              {Data => $self->{ParseOptions}->{IndentChar}}
      );
      $self->{ParseOptions}->{Handler}->end_element($el);
      $self->{ParseOptions}->{Handler}->characters({Data => $self->{ParseOptions}->{NewLine}});
  
  $self->{'_row'} = [];  ### Clear $row and start the new row processing
  
}

sub _print_xml_finish
{
  my $self = shift;
  
  while (@{$self->{'_row'}} < 9)
  {
    push(@{$self->{'_row'}}, undef);
  }
  
  my $el = {
        Name => $self->{ParseOptions}->{Parent_Tag} || "record",
        Attributes => {},
      };
      
      $self->{ParseOptions}->{Handler}->characters(
              {Data => $self->{ParseOptions}->{IndentChar} || "\t"
              }
      );
      $self->{ParseOptions}->{Handler}->start_element($el);
      $self->{ParseOptions}->{Handler}->characters({Data => $self->{ParseOptions}->{NewLine}});

      for (my $i = 0; $i <= $#{$self->{ParseOptions}->{Col_Headings}}; $i++) {
          my $column = { Name => $self->{ParseOptions}->{Col_Headings}->[$i], Attributes => {} };
          $self->{ParseOptions}->{Handler}->characters(
                  {Data => $self->{ParseOptions}->{IndentChar} x 2} 
          );
          $self->{ParseOptions}->{Handler}->start_element($column);
          $self->{ParseOptions}->{Handler}->characters({Data => $self->{'_row'}->[$i]});
          $self->{ParseOptions}->{Handler}->end_element($column);
          $self->{ParseOptions}->{Handler}->characters({Data => $self->{ParseOptions}->{NewLine}});
      }

      $self->{ParseOptions}->{Handler}->characters(
              {Data => $self->{ParseOptions}->{IndentChar}}
      );
      $self->{ParseOptions}->{Handler}->end_element($el);
      $self->{ParseOptions}->{Handler}->characters({Data => $self->{ParseOptions}->{NewLine}});
  
 
}


1;
__END__




=head1 NAME

  XML::SAXDriver::Exce; - SAXDriver for converting Excel files to XML

=head1 SYNOPSIS

    use XML::SAXDriver::Excel;
    my $driver = XML::SAXDriver::Excel->new(%attr);
    $driver->parse(%attr);

=head1 DESCRIPTION

  XML::SAXDriver::Excel was developed as a complement to 
  XML::Excel, though it provides a SAX interface, for 
  gained performance and efficiency, to Excel files.  
  Specific object attributes and handlers are set to 
  define the behavior of the parse() method.  It does 
  not matter where you define your attributes.  If they 
  are defined in the new() method, they will apply to 
  all parse() calls.  You can override in any call to 
  parse() and it will remain local to that function call 
  and not effect the rest of the object.

=head1 XML::SAXDriver::CSV properties

  Source - (Reference to a String, ByteStream, SystemId)
  
    String - Contains literal CSV data. 
             Ex (Source => {String => $foo})
      
    ByteStream - Contains a filehandle reference.  
                 Ex. (Source => {ByteStream => \*STDIN})
      
    SystemId - Contains the path to the file containing 
               the CSV data. Ex (Source => {SystemId => '../csv/foo.csv'})
      
  
  Handler - Contains the object to be used as a XML print handler
  
  DTDHandler - Contains the object to be used as a XML DTD handler.  
               ****There is no DTD support available at this time.  
               I'll make it available in the next version.****
  
  NewLine - Specifies the new line character to be used for printing XML data (if any).
            Defaults to '\n' but can be changed.  If you don't want to indent use empty 
            quotes.  Ex. (NewLine => "")
            
  IndentChar - Specifies the indentation character to be used for printing XML data (if any).
               Defaults to '\t' but can be changed.  Ex. (IndentChar => "\t\t")
               
  Col_Headings - Reference to the array of column names to be used for XML tag names.
  
  Dynamic_Col_Headings - Should be set if you want the XML tag names 
                         generated dynamically from the row in CSV 
                         file.
                         
=head1 AUTHOR

Ilya Sterin (isterin@cpan.org)

=head1 SEE ALSO

XML::Excel
Spreadsheet::ParseExcel

=cut
