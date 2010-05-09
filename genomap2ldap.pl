#!/usr/bin/perl

# Tested on WindowsXP, GenoPro v2.0.1.6, ActivePerl v5.10.0 (XML::Twig v3.32, Net::LDAP v0.39)

use strict;
use utf8;

use Encode;

use Carp;

use Archive::Zip qw(:ERROR_CODES :CONSTANTS);

use XML::Twig;
use Date::Calc;
use Net::LDAP;
use Net::LDAP::LDIF;
use Net::LDAP::Entry;
use File::Basename;

use Getopt::Long qw(:config bundling);
use Pod::Usage;

# Used for debugging:
#use Data::Dumper;

binmode(STDOUT, ':utf8');
binmode(STDERR, ':utf8');

# Command-line arguments:
my ($genomap_file, $ldap_host, $ldap_search_dn, $ldap_search_group_dn, $ldap_bind_dn, $ldap_bind_pass);

pod2usage(-verbose=>0, -exitval=>1) unless GetOptions(
	'help'					=> sub { pod2usage(-verbose=>2, -noperldoc=>1, -exitval=>1); },
	'file|f=s'				=> \$genomap_file,
	'host|h=s'				=> \$ldap_host,
	'search-dn|S=s'			=> \$ldap_search_dn,
	'search-group-dn|G=s'	=> \$ldap_search_group_dn,
	'bind-dn|D=s'			=> \$ldap_bind_dn,
	'bind-pass|w=s'			=> \$ldap_bind_pass,
) && defined $genomap_file;

# The list of all individuals (key = individualID)
my %individuals = ();

# The list of all contact informations (key = contactID)
my %contacts = ();

# The list of all pictures (key = pictureID)
my %pictures = ();

# The list of all places (key = placeID)
my %places = ();

# The list of all groups (with group name as a key)
my %groups = ();

# The list of all DNs in this session
my %dn = ();

######################################
# Helper functions
######################################

sub check_dn($)
{
	local $_ = shift;

	if (exists $dn{$_})
	{
		croak('DN ' . $_ . 'already defined in this session. Program cannot process the data reliably, as DN definition mechanizm should be changed.');
	}

	$dn{$_} = 1;
}

sub get_contact_type_weight($)
{
	local $_ = shift;

	return 5 unless defined and length > 0;
	return 1 if $_ eq 'PrimaryResidence';
	return 2 if $_ eq 'Other';
	return 3 if $_ eq 'TemporaryResidence';
	return 4 if $_ eq 'WorkPlace';

	croak "The program should never reach here for value '$_'";
}

sub trim($)
{
	local $_ = shift;
	s/^\s*//;
	s/\s*$//;
	return $_; 
}

######################################
# Document loading. XML Parsing
######################################

my $genomap_zip = Archive::Zip->new();
die "Unable to read ZIP file $genomap_file" unless $genomap_zip->read($genomap_file) == AZ_OK;
die "ZIP file is empty" unless $genomap_zip->numberOfMembers();

XML::Twig->new(
	twig_handlers => {
		'Individuals/Individual'	=> \&individual,
		'Contacts/Contact'			=> \&contact,
		'Pictures/Picture'			=> \&picture,
		'Places/Place'				=> \&place
	}
)->parse(($genomap_zip->members())[0]->contents())->purge();

while (local (undef, $_) = each %individuals)
{
	# Sort all contact information by relevance:
	my @contacts = sort {
		get_contact_type_weight($a->{'type'}) <=> get_contact_type_weight($b->{'type'})
	} map {
		$contacts{$_} or croak "No contacts found for ID $_";
	} @{$_->{'contacts'}};

	$_->{'contacts'} = \@contacts;

	# Define the group for individual:
	foreach my $group (@{$_->{'groups'}})
	{
		push @{$groups{$group}}, $_;
	}
	
	# Resolve the picture for individual:
	if (defined $_->{'picture'})
	{
		$_->{'picture'} = $pictures{$_->{'picture'}};
	}
}

# Merge the information from parent places:
while (my ($place_id, $place) = each %places)
{
	my $parent_place_id = $place->{'parent'};

	while (defined $parent_place_id)
	{
		if (!defined $places{$parent_place_id})
		{
			carp "Unable to locate parent place for '$parent_place_id'";
			last;
		}

		while (my ($property_key, $property_value) = each %{$places{$parent_place_id}})
		{
			if (!defined $place->{$property_key})
			{
				$place->{$property_key} = $property_value;
			}
		}

		$parent_place_id = $places{$parent_place_id}->{'parent'};
	}
}

######################################
# Loading to LDAP
######################################

my $ldap = undef;
my $ldif = Net::LDAP::LDIF->new(\*STDOUT, 'w', 'onerror' => 'die', 'encode' => 'base64'); # 'change' => 1

if (defined $ldap_host)
{
	$ldap = Net::LDAP->new($ldap_host, 'onerror' => 'die') or croak "$@";
	#$ldap->debug(15);

	print "Connected to host $ldap_host, LDAP v" . $ldap->version() . "\n";

	# Anonymous bind?
	if (defined $ldap_bind_dn && defined $ldap_bind_pass)
	{
		$ldap->bind($ldap_bind_dn, 'password' => $ldap_bind_pass);
	}
	else
	{
		$ldap->bind();
	}
}

while (my ($individual_id, $individual) = each %individuals)
{
	my $entry;

	# Contact information is already sorted by relevance:	
	foreach my $contact (@{$individual->{'contacts'}})
	{
		if (defined $contact->{email})
		{
			unless (defined $entry)
			{
				if ($contact->{email} !~ /^([-'\w]+)\s+([-'\w. ]+)\s+<(.*)>$/)
				{
					carp "Unable to parse email $contact->{email} for individual '$individual_id'";
					next;
				}

				my $dn = "cn=$1 $2" . (defined $ldap_search_dn ? ',' . $ldap_search_dn : '');
				
				eval {
					$entry = $ldap->search('base' => $dn, 'filter' => '(objectClass=person)', 'scope' => 'base', 'sizelimit' => 1)->pop_entry()
				} if $ldap;

				# DN was not found or we have no LDAP connection:
				unless (defined $entry)
				{
					$entry = Net::LDAP::Entry->new($dn);

					# These attributes only for adding a new entry:
					$entry->add(
						'givenName' => $1,
						'sn' => $2,
						'cn' => "$1 $2",
						'objectClass' => [ qw(inetOrgPerson mozillaAbPersonAlpha) ]
					);
				}

				$individual->{'dn'} = $dn;

				check_dn($dn);

				# Other attributes for adding or updating the entry:
				$entry->replace('mail' => encode('MIME-Q', $3));
			}
			else
			{
				if ($contact->{email} !~ /<(.*)>$/)
				{
					carp "Unable to parse email $contact->{email} for individual $individual_id";
					next;
				}

				$entry->replace('mozillaSecondEmail' => encode('MIME-Q', $1)) if !$entry->exists('mozillaSecondEmail');
				
				if ($entry->get_value('mozillaSecondEmail') eq $entry->get_value('mail'))
				{
					$entry->delete('mozillaSecondEmail' => undef);
				}
			}
		}
	}

	unless (defined $entry)
	{
		#carp "Individual $individual_id has no email";
		next;
	}

	# Processing other attributes after the entry has been created:

	foreach my $contact (@{$individual->{'contacts'}})
	{
		$entry->replace('telephoneNumber'	=> $contact->{'telephone'})	if defined $contact->{'telephone'}	&& !$entry->exists('telephoneNumber');
		$entry->replace('mobile'			=> $contact->{'mobile'})	if defined $contact->{'mobile'}		&& !$entry->exists('mobile');
		$entry->replace('mozillaHomeUrl'	=> $contact->{'homepage'})	if defined $contact->{'homepage'}	&& !$entry->exists('mozillaHomeUrl');
		
		if (defined $contact->{'place'})
		{
			if ($contact->{'type'} eq 'WorkPlace')
			{
				$entry->replace('street'					=> $contact->{'place'}->{'street'})		if defined $contact->{'place'}->{'street'}	&& !$entry->exists('street');
				$entry->replace('postalCode'				=> $contact->{'place'}->{'zip'})		if defined $contact->{'place'}->{'zip'}		&& !$entry->exists('postalCode');
				$entry->replace('l'							=> $contact->{'place'}->{'city'})		if defined $contact->{'place'}->{'city'}	&& !$entry->exists('l');
				$entry->replace('c'							=> $contact->{'place'}->{'country'})	if defined $contact->{'place'}->{'country'}	&& !$entry->exists('c');
			}
			else
			{
				$entry->replace('mozillaHomeStreet'			=> $contact->{'place'}->{'street'})		if defined $contact->{'place'}->{'street'}	&& !$entry->exists('mozillaHomeStreet');
				$entry->replace('mozillaHomePostalCode'		=> $contact->{'place'}->{'zip'})		if defined $contact->{'place'}->{'zip'}		&& !$entry->exists('mozillaHomePostalCode');
				$entry->replace('mozillaHomeLocalityName'	=> $contact->{'place'}->{'city'})		if defined $contact->{'place'}->{'city'}	&& !$entry->exists('mozillaHomeLocalityName');
				$entry->replace('mozillaHomeCountryName'	=> $contact->{'place'}->{'country'})	if defined $contact->{'place'}->{'country'}	&& !$entry->exists('mozillaHomeCountryName');
			}
		}
	}

	$entry->replace('pager'					=> $individual->{'icq'})	if defined $individual->{'icq'};
	
	if (defined $individual->{'birth_date'})
	{
		if ($individual->{'birth_date'} =~ /(\d{1,2}) (\w{3,3})(?: (\d{4,4}))?/)
		{
			$entry->replace('birthyear'		=> $3) if defined $3;
			$entry->replace('birthmonth'	=> Date::Calc::Decode_Month($2));
			$entry->replace('birthday'		=> $1);
		}
	}

	if (defined $individual->{'picture'})
	{
		local $/ = undef;
		local $_ = $individual->{'picture'};
		s#\\#\/#g;
		open IN, "<", (fileparse($genomap_file))[1] . encode("cp1251", $_) or die $!;
		$entry->replace('jpegPhoto' => <IN>);
		close IN;
	}

	if ($ldap)
	{
		print(($entry->changetype() eq 'add' ? "Adding " : "Updating ") . $entry->dn() . "\n");
		$entry->update($ldap);
	}
	else
	{
		$ldif->write_entry($entry);
	}
}

while (my ($cn, $group) = each %groups)
{
	my $entry;
	my $dn = 'cn=' . $cn . (defined $ldap_search_group_dn ? ',' . $ldap_search_group_dn : '');

	eval {
		$entry = $ldap->search('base' => $dn, 'filter' => '(objectClass=groupOfNames)', 'scope' => 'base', 'sizelimit' => 1)->pop_entry()
	} if $ldap;

	# DN was not found or we have no LDAP connection:
	unless (defined $entry)
	{ 
		$entry = Net::LDAP::Entry->new($dn);
	
		$entry->add(
			'objectclass' => 'groupOfNames',
			'cn' => $cn
		);
	}
	
	my %saw;
	$entry->replace('member' => [ map { $_->{'dn'} } grep { defined $_->{'dn'} && !$saw{$_->{'dn'}}++ } @{$group} ]) ;
	
	# No members for this group have been added to LDAP:
	next unless scalar($entry->get_value('member'));
	
	check_dn($dn);
	
	if ($ldap)
	{
		print(($entry->changetype() eq 'add' ? "Adding" : "Updating") . " group: $dn\n");
	
		$entry->update($ldap);
	}
	else
	{
		$ldif->write_entry($entry);
	}
}

$ldap->unbind() if $ldap;

######################################
# XML callback functions
######################################

sub individual()
{
	my ($twig, $individual_node) = @_;
	
	my $map_name = $individual_node->first_child('Position')->att('GenoMap');
	my $individual_link_id = $individual_node->att('IndividualInternalHyperlink');
	
	# If this individual is actually a link, we add a map name to the list of groups.
	# It is assumed, that linked individuals go in XML later than individuals they refer to
	# (e.g. $individual_link_id > $individual_id).
	if (defined $individual_link_id)
	{
		push @{$individuals{$individual_link_id}->{'groups'}}, $map_name if defined $map_name;
		return;	
	}
	
	my $individual_id = $individual_node->att('ID');
	my $contacts_node = $individual_node->first_child('Contacts');

	if (!defined $contacts_node)
	{
		#carp('No contact information available for individual ' . $individual_id); 
		return;
	}

	my @contacts = split /\s*,\s*/, $contacts_node->text();

	if (scalar(@contacts) == 0)
	{
		carp('Contact information list is empty for individual ' . $individual_id);
		return;
	}

	local $_ = {};
	$individuals{$individual_id} = $_;
	$_->{'contacts'} = \@contacts;

	if (defined $map_name)
	{
		$_->{'groups'} = [ $map_name ]; 
	}

	# Custom tags:
	foreach my $node_path (qw(Name ICQ Skype JabberId AIM Birth/Date))
	{
		my $property_value = $individual_node->findvalue($node_path);

		my $property_name = $node_path;
		$property_name =~ s/\//_/g;
		
		$_->{lc $property_name} = trim($property_value) if $property_value;
	}
	
	my $pictures_node = $individual_node->first_child('Pictures');
	
	# As pictures are parsed after individuals, we have to save the picture ID to resolve later:
	if (defined $pictures_node)
	{
		$_->{'picture'} = $pictures_node->att('Primary');
	}
}

sub contact()
{
	my ($twig, $contact_node) = @_;

	my %properties = ();

	$contacts{$contact_node->att('ID')} = \%properties;

	foreach (qw(Type Email Telephone Mobile Homepage))
	{
		my $property_value = $contact_node->findvalue($_);
		$properties{lc $_} = trim($property_value) if $property_value;
	}

	my $place_id = $contact_node->findvalue('Place');
	$properties{'place'} = $places{$place_id} if defined $places{$place_id};
}

sub picture()
{
	my ($twig, $picture_node) = @_;

	$pictures{$picture_node->att('ID')} = $picture_node->first_child('Path')->att('Relative');
}

sub place()
{
	my ($twig, $place_node) = @_;

	my %properties = ();

	$places{$place_node->att('ID')} = \%properties;

	foreach (qw(Name Country City Street Zip Parent))
	{
		my $property_value = $place_node->findvalue($_);
		$properties{lc $_} = trim($property_value) if $property_value;
	}
}

__END__
=pod

=head1 NAME

genomap2ldap.pl - loads infromation about GenoMap individuals into LDAP directory

=head1 SYNOPSIS

  genomap2ldap.pl -f genomap.gno [-h ldap.host.com] [-D bind_dn] [-w bind_password] [-S users_dn] [-G groups_dn]
  genomap2ldap.pl --help

  genomap2ldap.pl -f C:/Documents/myfile.gno -S cn=user,cn=domain -G cn=groups,cn=domain -h 192.168.1.5 -D cn=ldapadmin,cn=domain -w superUserPass

=head1 OPTIONS

=over

=item B<--help>

Prints this help message.

=item B<--host|-h>

Optionally specify the LDAP host to connect to. If not defined, the data is printed to standard output in LDIFF format,
which can be manually imported into LDAP directory.

=item B<--search-dn|-S>

Optionally specify the search DN, which should be used during the lookups for individuals. If this is not defined,
all lookups are made in root DN ("").

=item B<--search-group-dn|-G>

Optionally specify the search DN, which should be used during the lookups for groups. If this is not defined,
all lookups are made in root DN ("").

=item B<--bind-dn|-D>

Optionally specify the bind DN, which should be used during the authentication phase. Note, that both bind DN and bind password
should be specified, otherwise the anonymous bind will take place.

=item B<--bind-pass|-w>

Optionally specify the bind password, which should be used during the authentication phase.

=back

=head1 DESCRIPTION

The program opens map file using GenoMap COM interface. The XML data is parsed and converted to .ldiff format.
The program prints the resulting .ldiff to the screen, if no LDAP host is provided. Otherwise for each added
individual it is checked, if the entry already exists in LDAP directory. If positive, the entry fields are
updated, otherwise a new entry is added.

=head1 LDAP Attributes Used by Thunderbird

birthday o company mail modifytimestamp mozillaUseHtmlMail xmozillausehtmlmail mozillaCustom2 custom2
mozillaHomeCountryName ou department departmentnumber orgunit mobile cellphone carphone telephoneNumber title
mozillaCustom1 custom1 sn surname mozillaNickname xmozillanickname mozillaWorkUrl workurl labeledURI
facsimiletelephonenumber fax mozillaSecondEmail xmozillasecondemail mozillaCustom4 custom4 nsAIMid nscpaimscreenname
street streetaddress postOfficeBox givenName l locality homePhone mozillaHomeUrl homeurl mozillaHomeStreet st region
mozillaHomePostalCode mozillaHomeLocalityName mozillaCustom3 custom3 birthyear mozillaWorkStreet2 mozillaHomeStreet2
postalCode zip birthmonth c countryname pager pagerphone mozillaHomeState description notes cn commonname objectClass

More information:
http://wiki.mozilla.org/MailNews:Mozilla%20LDAP%20Address%20Book%20Schema
http://www.mozilla.org/projects/thunderbird/specs/ldap.html

=cut
