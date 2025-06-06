import os
import argparse
import pandas as pd
from Bio import SeqIO, Seq, SeqRecord
from Bio.Align import PairwiseAligner
from Bio.Data import CodonTable
import glob

def parse_ab1_file(file_path, trim_start=50):
    """Parse AB1 file and return the trimmed sequence."""
    try:
        record = SeqIO.read(file_path, "abi")
        # Trim the first 50 nucleotides
        trimmed_seq = record.seq[trim_start:]
        return str(trimmed_seq)
    except Exception as e:
        print(f"Error parsing {file_path}: {e}")
        return None

def detect_orientation(seq, ref_seq, aligner):
    """Detect if sequence is forward or reverse complement."""
    
    forward_score = aligner.score(seq, ref_seq)
    reverse_seq = str(Seq.Seq(seq).reverse_complement())
    reverse_score = aligner.score(reverse_seq, ref_seq)
    
    if reverse_score > forward_score:
        return "reverse", reverse_seq
    else:
        return "forward", seq

def align_sequence(seq, ref_seq, min_aligned_length=50):
    """Align sequence to reference sequence."""
    aligner = PairwiseAligner()
    aligner.mode = 'global'
    # Medium stringency settings
    aligner.match_score = 2
    aligner.mismatch_score = -1
    aligner.open_gap_score = -2
    aligner.extend_gap_score = -0.5
    
    # Detect orientation
    orientation, oriented_seq = detect_orientation(seq, ref_seq, aligner)
    
    # Perform alignment
    alignments = aligner.align(oriented_seq, ref_seq)
    best_alignment = alignments[0]
    
    return best_alignment, orientation

def find_mutations(alignment, ref_seq):
    """Find mutations between aligned sequence and reference."""
    aligned_seq, aligned_ref = alignment[0], alignment[1]
    mutations = []
    
    seq_pos = 0
    ref_pos = 0
    
    for i in range(len(aligned_seq)):
        seq_char = aligned_seq[i]
        ref_char = aligned_ref[i]
        
        if seq_char != '-':
            seq_pos += 1
        
        if ref_char != '-':
            ref_pos += 1
            
        if seq_char != ref_char and seq_char != '-' and ref_char != '-':
            mutations.append({
                'ref_pos': ref_pos,
                'ref_base': ref_char,
                'seq_base': seq_char
            })
    
    return mutations

def analyze_codon_changes(mutations, ref_seq):
    """Analyze codon changes and amino acid impacts."""
    # Group mutations by codon position
    codon_mutations = {}
    
    for mut in mutations:
        ref_pos = mut['ref_pos']
        codon_pos = (ref_pos - 1) // 3
        if codon_pos not in codon_mutations:
            codon_mutations[codon_pos] = []
        codon_mutations[codon_pos].append(mut)
    
    # Analyze each affected codon
    standard_table = CodonTable.standard_dna_table
    mutation_results = []
    
    for codon_pos, muts in codon_mutations.items():
        # Get original codon
        start_pos = codon_pos * 3
        orig_codon = ref_seq[start_pos:start_pos+3]
        
        if len(orig_codon) < 3:
            continue  # Skip incomplete codons
        
        # Create mutated codon
        mutated_codon = list(orig_codon)
        for mut in muts:
            codon_index = (mut['ref_pos'] - 1) % 3
            mutated_codon[codon_index] = mut['seq_base']
        
        mutated_codon = ''.join(mutated_codon)
        
        # Translate codons to amino acids
        try:
            orig_aa = standard_table.forward_table.get(orig_codon, "X")
            mutated_aa = standard_table.forward_table.get(mutated_codon, "X")
            
            is_silent = (orig_aa == mutated_aa)
            
            mutation_results.append({
                'codon_position': codon_pos + 1,  # 1-based position
                'nucleotide_position': start_pos + 1,  # 1-based position
                'original_codon': orig_codon,
                'mutated_codon': mutated_codon,
                'original_aa': orig_aa,
                'mutated_aa': mutated_aa,
                'aa_position': codon_pos + 1,  # 1-based position
                'is_silent': is_silent,
                'mutation_type': 'Silent' if is_silent else 'Missense'
            })
        except Exception as e:
            print(f"Error analyzing codon {orig_codon} -> {mutated_codon}: {e}")
    
    return mutation_results

def main():
    parser = argparse.ArgumentParser(description='Sequence Aligner for AB1 files')
    parser.add_argument('--ref', required=True, help='Reference sequence file (FASTA)')
    parser.add_argument('--input', required=True, help='Directory containing AB1 files')
    parser.add_argument('--output', default='mutations.xlsx', help='Output Excel file')
    parser.add_argument('--trim', type=int, default=50, help='Number of nucleotides to trim from start')
    parser.add_argument('--min_length', type=int, default=50, help='Minimum aligned read length')
    
    args = parser.parse_args()
    
    # Load reference sequence
    ref_record = SeqIO.read(args.ref, "fasta")
    ref_seq = str(ref_record.seq)
    
    # Find all AB1 files
    ab1_files = glob.glob(os.path.join(args.input, "*.ab1"))
    
    if not ab1_files:
        print(f"No AB1 files found in {args.input}")
        return
    
    all_results = []
    
    for ab1_file in ab1_files:
        sample_name = os.path.basename(ab1_file).split('.')[0]
        print(f"Processing {sample_name}...")
        
        # Parse AB1 file
        seq = parse_ab1_file(ab1_file, args.trim)
        if not seq:
            continue
            
        # Align sequence
        alignment, orientation = align_sequence(seq, ref_seq, args.min_length)
        
        # Find mutations
        mutations = find_mutations(alignment, ref_seq)
        
        # Analyze codon changes
        codon_results = analyze_codon_changes(mutations, ref_seq)
        
        # Add sample information
        for result in codon_results:
            result['sample'] = sample_name
            result['orientation'] = orientation
            all_results.append(result)
    
    # Create DataFrame and export to Excel
    if all_results:
        df = pd.DataFrame(all_results)
        
        # Reorder columns for better readability
        columns = ['sample', 'orientation', 'nucleotide_position', 'original_codon', 
                  'mutated_codon', 'aa_position', 'original_aa', 'mutated_aa', 
                  'is_silent', 'mutation_type']
        
        df = df[columns]
        df.sort_values(['sample', 'nucleotide_position'], inplace=True)
        
        # Export to Excel
        df.to_excel(args.output, index=False)
        print(f"Results exported to {args.output}")
    else:
        print("No mutations found.")

if __name__ == "__main__":
    main()