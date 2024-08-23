import os
import numpy as np
from Bio import Entrez, SeqIO
from Bio.Blast import NCBIXML
from scipy import stats
import statsmodels.stats.multicomp as mc
import pandas as pd
import subprocess
import shutil
import time


#Things that are important throughout
folder_path ='/home/srinivas/ArchaeaGenomeComparer'
Entrez.email = 'sriedu101@gmail.com'
Entrez.api_key = '66ef73a2c3efa4447b7f3e4ef38d8188fe09'

class Expierement:
    def __init__(self,organisms:np.array,proteins:np.array,sigLevel):
        self.organisms = organisms
        self.proteins = proteins
        self.sigLevel = sigLevel


    def organismsDownloader (self):
        organismSummary = f'{folder_path}/organismSummary.txt'
        finalizedOrganismList = np.array([])
        for organism in self.organisms:
            # Initial Pause, Do not want to overload servers, 5 minutes each
            print (f'Will Download {organism} genome in 5 minutes')
            time.sleep(300)
            
            # Clean up string
            splicedName = organism.split()
            organismType = "".join(splicedName)

            # Download and Unzip files
            command = ['/home/srinivas/anaconda3/envs/ncbi_datasets/bin/datasets', 'download', 'genome', 'taxon', f'{organism}', '--include', 'genome', '--filename',f'{organismType}.zip']
            subprocess.run(command,capture_output=True, text=True).stdout
            unzipcommand = ['/usr/bin/unzip',f'{organismType}.zip']
            subprocess.run(unzipcommand,capture_output=True, text=True).stdout

            # Check to see if there is a file
            zipchecker = os.listdir(folder_path)
            if f'{organismType}.zip' in zipchecker:
                # Set Up directories to hold things in the end and the blast database file
                directory = f'{organismType}'
                parent_dir = folder_path
                path = os.path.join(parent_dir, directory)
                os.mkdir(path)
                specific_folder_path = f'{parent_dir}/{directory}'
                trash_folder_path = f'{parent_dir}/Trash'
                blastDatabaseFileName = f"{parent_dir}/BlastDatabase_{organismType}.fasta"

                # Begin copy pasting the details from NCBIs Genome Dataset into the single Blast Database file
                with open(blastDatabaseFileName,'w') as output_file:
                    counter = 1 
                    for folder_name in os.listdir(f'{parent_dir}/ncbi_dataset/data'):
                        inside_folder_path = os.path.join(f'{parent_dir}/ncbi_dataset/data', folder_name)
                        if os.path.isdir(inside_folder_path):
                            files = os.listdir(inside_folder_path)
                            for file in files:
                                file_location = f'{inside_folder_path}/{file}'
                                with open(file_location, 'r') as fna_file:
                                    file_contents = fna_file.read()
                                    output_file.write (f'>{organismType} Seq{counter}')
                                    output_file.write(file_contents)
                                    output_file.write("\n")
                                    counter = counter+1
                        else:
                            print(f"{inside_folder_path} is not a directory.")

                # Have the summary of everything that was saved onto a secondary document            
                with open(organismSummary,'a') as output_file:
                    with open(f'{parent_dir}/ncbi_dataset/data/assembly_data_report.jsonl','r') as summary_file:
                        file_contents = summary_file.read()
                        output_file.write(file_contents)

                # Now that every relvent information is saved and database file is made, run make blast datbase and make it into a database
                makeblastcommand = ['makeblastdb', '-in', f'{blastDatabaseFileName}', '-input_type', 'fasta', '-dbtype', 'nucl', '-out', f'BlastDatabase_{organismType}']
                subprocess.run(makeblastcommand,capture_output=True, text=True).stdout

                # Now that we have everything we need lets move the folders to their apprpiate place
                shutil.move(f'{folder_path}/ncbi_dataset',f'{specific_folder_path}/ncbi_dataset') # Move DataSets to their specific folder
                shutil.move(f'{folder_path}/README.md',f'{specific_folder_path}/README.md') # Move the ReadMe details to the specific folder again
                shutil.move(f'{specific_folder_path}',f'{folder_path}/Storage/{directory}') # Then move that full folder into storage
                shutil.move(f'{folder_path}/{organismType}.zip',f'{trash_folder_path}/{organismType}.zip') # Move the zip into a trash folder that can be deleted in the end

                # Last thing to do is update the existing np array
                finalizedOrganismList = np.append(finalizedOrganismList,f'{organism}')
            else:
                print (f'{organism} genomes not found will skip.')
            
        # New updated organism list with only organisms that have been downloaded and set up
        self.organisms = finalizedOrganismList

    def proteinDownloader(self):
        placement = 1 # Initial protein index will start with one
        tempProteinDataSummary = pd.DataFrame(columns=["Protein Name","Protein ID","Protein Title", "Additional Protein Information"]) # Protein Summaary to store
        proteins = self.proteins
        for protein in proteins:
            # Inital Pause for 10 minutes just to not overload the servers
            print (f'Will download {protein} in 5 minutes...')
            time.sleep(300)

            # Use Entrez to get the first protein sequence for the given protein
            search_term = f'{protein}[Protein Name] AND Saccharomyces Cerevisiae [Organism Name]'
            handle = Entrez.esearch(db='protein', term=search_term, retmax=1)
            record = Entrez.read(handle)

            # Make Sure we have a protein to work with if not return
            if int(record["Count"]) != 0:
                # Retrieve the ID for the matching protein
                protein_id = record['IdList'][0]

                # Fetch the protein sequence from NCBI
                handle = Entrez.efetch(db='protein', id=protein_id, rettype='fasta', retmode='text')
                record = SeqIO.read(handle, 'fasta')

                # Save the protein sequence to a file
                filename = f'Protein_{placement}.fasta'
                with open(filename, 'w') as file:
                    SeqIO.write(record, file, 'fasta')
                
                # Update Table Accordingly
                summaryHandle = Entrez.esummary(db='protein', id=protein_id, rettype='fasta')
                summary = Entrez.read(summaryHandle)
                proteinInformation = np.array([])
                proteinInformation = np.append(proteinInformation,protein)
                proteinInformation = np.append(proteinInformation,protein_id)
                proteinInformation = np.append(proteinInformation,summary[0]['Title'])
                proteinInformation = np.append(proteinInformation,summary[0]['Extra'])
                tempProteinDataSummary.loc[len(tempProteinDataSummary)]=proteinInformation

                #Increment Placement by 1 moving to the next placement
                placement=placement+1
                print(f'Protein sequence for {protein} downloaded and saved successfully.')
            else:
                print(f"No protein records found for {protein}")
        # Save the summary
        with open(f'{folder_path}/proteinSummary.txt', 'w') as txt_file:
            txt_file.write(tempProteinDataSummary.to_string(index=False))


    def assessSimilarity(self,counter):
        # Folder setup for clean view
        directory = f'protein{counter}_result'
        parent_dir = folder_path
        path = os.path.join(parent_dir, directory)
        os.mkdir(path)
        specific_folder_path = f'{parent_dir}/{directory}'

        # Array set up to hold data
        averageScores = pd.DataFrame(columns=["Protein Name","Organism ID","BLAST Score Average", "BIT Score Average", "EScore Average"])
        listofBLASTScores = []
        listofBITScores = []
        listofEScores = []

        # Proceed to go through each organism to aquire the needed results
        organisms = self.organisms
        for organism in organisms:
            # Clean up string
            splicedName = organism.split()
            organismType = "".join(splicedName)

            # This is the internal list, I am keeping this seperate from the other as this is a list of numbers for each organism while the other holds all values
            listofBLASTScoresInternal = np.array([],dtype=np.int64)
            listofBITScoresInternal = np.array([],dtype=np.int64)
            listofEScoresInternal = np.array([],dtype=np.int64)

            # Run the Blast Commnad and store the value in an xml file, section works
            command = ['tblastn', '-query', f'Protein_{counter}.fasta', '-db', f'BlastDatabase_{organismType}', '-out', f'{organismType}_protein{counter}.out', '-outfmt', '5']
            subprocess.run(command, capture_output=True, text=True).stdout
            with open(f'{organismType}_protein{counter}.out', 'r') as out_file:
                result = out_file.read()
            with open(f'{organismType}_protein{counter}_result.xml', 'w') as xml_file:
                xml_file.write(result)
            file_out = f'{organismType}_protein{counter}.out'
            file_xml = f'{organismType}_protein{counter}_result.xml'
            if os.stat(f'{parent_dir}/{file_xml}').st_size!=0:
                #If it is not empty parse it
                blastFile = NCBIXML.parse(open(f'{organismType}_protein{counter}_result.xml'))
                for record in blastFile:
                    for alignment in record.alignments:
                        for hsp in alignment.hsps:
                            if hsp.bits>40: # Thereshold for score
                                # Store the score, bit score (score but standardized), and escore (level of surrounding noise)
                                score = hsp.score  
                                bits = hsp.bits
                                escore = hsp.expect 
                                listofBLASTScoresInternal = np.append(listofBLASTScoresInternal,score)
                                listofBITScoresInternal = np.append(listofBITScoresInternal,bits)
                                listofEScoresInternal = np.append(listofEScoresInternal,escore)
            
            # For each of these values, we want to collect and store the average for each value. This works as an if else in the case there were no hits
            if len(listofBLASTScoresInternal)!=0:
                averageBLASTScore = sum(listofBLASTScoresInternal)/len(listofBLASTScoresInternal)
                listofBLASTScores.append(listofBLASTScoresInternal)
            else: 
                averageBLASTScore = 0
                listofBLASTScoresInternal = np.append(listofBLASTScoresInternal,0)
                listofBLASTScores.append(listofBLASTScoresInternal)

            if len(listofBITScoresInternal) != 0:
                averageBITScore = sum(listofBITScoresInternal)/len(listofBITScoresInternal)
                listofBITScores.append(listofBITScoresInternal)
            else: 
                averageBITScore = 0
                listofBITScoresInternal = np.append(listofBITScoresInternal,0)
                listofBITScores.append(listofBITScoresInternal)

            if len(listofEScoresInternal) != 0: 
                averageEScore = sum(listofEScoresInternal)/len(listofEScoresInternal)
                listofEScores.append(listofEScoresInternal)
            else: 
                averageEScore = 0
                listofEScoresInternal = np.append(listofEScoresInternal,0)
                listofEScores.append(listofEScoresInternal)

            # Store and Save said values into the Average Table 
            addToAverageTable = np.array([])
            addToAverageTable = np.append(addToAverageTable,counter)
            addToAverageTable = np.append(addToAverageTable,organismType)
            addToAverageTable = np.append(addToAverageTable,averageBLASTScore)
            addToAverageTable = np.append(addToAverageTable,averageBITScore)
            addToAverageTable = np.append(addToAverageTable,averageEScore)
            averageScores.loc[len(averageScores)]=addToAverageTable

            # Before moving to the enxt organism, move these files over to the results folder for organism
            shutil.move(f'{folder_path}/{file_out}',f'{specific_folder_path}/{file_out}')
            shutil.move(f'{folder_path}/{file_xml}',f'{specific_folder_path}/{file_xml}')

        # Saving Averages
        with open(f'{parent_dir}/averageValueSummary.txt', 'a') as txt_file:  
            txt_file.write(averageScores.to_string(index=False))

        # Saving Raw Values
        with open(f'{parent_dir}/rawValueSummary.txt', 'a') as txt_file:   
            # First Tell what protein we are workin
            txt_file.write (f'Protein {counter} current information \n')
            index = 0 
            for organismGroup in organisms:
                txt_file.write (f'\n{organismGroup} current information \n')
                
                txt_file.write ("BLAST Scores\n")
                blastNumpyArray = np.array2string(listofBLASTScores[index])
                txt_file.write (blastNumpyArray)

                txt_file.write ("\n"+"BIT Scores \n")
                bitNumpyArray = np.array2string(listofBITScores[index])
                txt_file.write(bitNumpyArray)

                txt_file.write ("\n"+"E Scores \n")
                eNumpyArray = np.array2string(listofEScores[index])
                txt_file.write(eNumpyArray)
                index = index+1

            txt_file.write('. \n')
            txt_file.write('. \n')

        # Final move the protein results folder into the individual results folder
        shutil.move(f'{specific_folder_path}',f'{parent_dir}/individual_protein_results/protein{counter}_result')

        return listofBLASTScores
    


    def statisticalAnalysis(self,lists):
        _, p_value = stats.f_oneway(*lists)
        result_str = ""

        # If the p-value is less than the significance level
        if p_value < self.sigLevel:
            result_str += f'{p_value} is the p value. It is less than the significance level, {self.sigLevel}\n'
        else:
            result_str += f'{p_value} is the p value. It is not less than the significance level, {self.sigLevel}\n'
        post_hoc_result = self.postHoc(lists)
        result_str += str(post_hoc_result) + '\n'
        result_str += '_____________________________________________________________________________\n'
        
        return result_str



    def postHoc (self,data_list):
        groupNumbers = []
        for i, sublist in enumerate(data_list):
            groupNumbers.extend([str(i)] * len(sublist))
        flattened_data = [item for sublist in data_list for item in sublist]
        pairwise_comparisons = mc.MultiComparison(flattened_data, groupNumbers)
        result = pairwise_comparisons.tukeyhsd()
        return result.summary()



    def conduct (self):
        # First Download Organisms and Proteins
        # self.organismsDownloader() # Finished this part so now moving to the next one
        self.proteinDownloader()
        
        # Create Directory for storing individual results
        parent_dir = folder_path
        individual_result_directory = 'individual_protein_results'
        individual_pathway = os.path.join(parent_dir,individual_result_directory)
        os.mkdir(individual_pathway)
        # Proteins are essentially labeled by an index starting from 1, using that conduct the expierment on all proteins
        i = 1
        for protein in self.proteins:
            result = self.assessSimilarity(i)
            insert = self.statisticalAnalysis(result)
            with open(f'{folder_path}/StatisticResults.txt', 'a') as file:
                file.write(f'Statistical Results for Protein {i}:\n')
                file.write(insert)
                file.write(f'________________________________________________________\n')
            print (f'Expierment for {protein} is complete.')
            i = i + 1
        
        # Once the expierement is over, to make results be at a easy place for access, make a result folder and store all the information summary over there
        result_directory = 'Results'
        result_path = os.path.join(parent_dir, result_directory)
        os.mkdir(result_path)
        shutil.move(f'{folder_path}/organismSummary.txt',f'{result_path}/organismSummary.txt') 
        shutil.move(f'{folder_path}/proteinSummary.txt',f'{result_path}/proteinSummary.txt') 
        shutil.move(f'{folder_path}/rawValueSummary.txt',f'{result_path}/rawValueSummary.txt') 
        shutil.move(f'{folder_path}/averageValueSummary.txt',f'{result_path}/averageValueSummary.txt') 
        shutil.move(f'{folder_path}/StatisticResults.txt',f'{result_path}/StatisticResults.txt') 

# Settings: List of Organisms and Proteins should be entered here. Set signfiicant level appropriately as well. This is the template used for the project, use whatever
# groups needed for your version

organisms = np.array([])

#Cluster 1
#organisms = np.append(organisms, "Methanobacteria") # Methanomada  < = Not found
#organisms = np.append(organisms, "Candidatus Idunnarchaeota") # Asgard   < = Not found

organisms = np.append(organisms, "Methanococci") # Methanomada 1
organisms = np.append(organisms, "Methanopyri") # Methanomada 2

organisms = np.append(organisms, "Thermococci") # Acherontia 3
organisms = np.append(organisms, "Theionarchaea") # Acherontia 4

organisms = np.append(organisms, "Candidatus Hadarchaeia") # Stygia 5

organisms = np.append(organisms, "Candidatus Odinarchaeota") # Asgard 6 
organisms = np.append(organisms, "Candidatus Wukongarchaeota") # Asgard 7 
organisms = np.append(organisms, "Candidatus Freyarchaeota") # Asgard 8
organisms = np.append(organisms, "Candidatus Baldrarchaeota") # Asgard 9
organisms = np.append(organisms, "Candidatus Heimdallarchaeota") # Asgard 10
organisms = np.append(organisms, "Candidatus Helarchaeota") # Asgard 11
organisms = np.append(organisms, "Candidatus Kariarchaeota") # Asgard 12
organisms = np.append(organisms, "Candidatus Lokiarchaeota") # Asgard 13
organisms = np.append(organisms, "Candidatus Sifarchaeota") # Asgard 14
organisms = np.append(organisms, "Candidatus Thorarchaeota") # Asgard 15
organisms = np.append(organisms, "Candidatus Hodarchaeota") # Asgard 16
organisms = np.append(organisms, "Candidatus Hermodarchaeota") # Asgard 17

organisms = np.append(organisms, "Candidatus Brockarchaeota") # TACK 18
organisms = np.append(organisms, "Candidatus Culexarchaeota") # TACK 19
organisms = np.append(organisms, "Candidatus Geoarchaeota") # TACK 20
organisms = np.append(organisms, "Candidatus Geothermarchaeota") # TACK 21
organisms = np.append(organisms, "Candidatus Korarchaeota") # TACK 22
organisms = np.append(organisms, "Candidatus Marsarchaeota") # TACK 23
organisms = np.append(organisms, "Candidatus Nezhaarchaeota") # TACK 24
organisms = np.append(organisms, "Candidatus Verstraetearchaeota") # TACK 25
organisms = np.append(organisms, "Nitrososphaerota") # TACK 26
organisms = np.append(organisms, "Thermoproteota") # TACK 27

#Cluster 2
organisms = np.append(organisms, "Candidatus Thermoplasmatota") # Diaforarchaea 28

organisms = np.append(organisms, "Methanonatronarchaeia") # Methanonatronarchaeia 29

organisms = np.append(organisms, "Candidatus Methanofastidiosia") # Stenosarchaea 30
organisms = np.append(organisms, "Candidatus Nanohaloarchaea") # Stenosarchaea 31
organisms = np.append(organisms, "Halobacteria") # Stenosarchaea 32
organisms = np.append(organisms, "Methanomicrobia") # Stenosarchaea 33




proteins = np.array([])
proteins = np.append (proteins, 'Smd3') # Core snRNP Protein
proteins = np.append (proteins, 'Prp9')  # RNA Splicing Factor
proteins = np.append (proteins, 'Histone H2A') # DNA Associated
proteins = np.append (proteins, 'Histone H3') # DNA Associated 
proteins = np.append (proteins, 'Histone H4') # DNA Associated  
proteins = np.append (proteins, 'Histone H2B') # DNA Associated  
proteins = np.append (proteins, 'Lsm2') # DNA Associated  
proteins = np.append (proteins, 'Rex3') # DNA Associated  
proteins = np.append (proteins, 'Ceg1') # DNA Associated  
proteins = np.append (proteins, 'Abd1') # DNA Associated  




sigLevel = 0.05

print ('Beginning Expierement')
expiermentCreation = Expierement(organisms,proteins,sigLevel)
expiermentCreation.conduct()