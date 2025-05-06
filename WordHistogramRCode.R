
library(readtext)
library(officer) 
library(ggplot2)
library(dplyr)
library(stringr)
library(tm)
library(data.table)
# !!! MAKE SURE TO UPLOAD YOUR FILE BEFORE RUNNING !!!

# Function to extract text from a specific section of the document
extract_text_from_section <- function(file_path, start_keyword) {
  # Read the document
  doc <- read_docx(file_path)
  text <- tolower(doc$text)

  # Extract all paragraphs from the document
  paragraphs <- docx_summary(doc)$text
  
  # Find the index where the start keyword appears
  start_index <- which(str_detect(paragraphs, regex(start_keyword, ignore_case = TRUE)))
  
  # If the start keyword is found, extract text from that point onward
  if (length(start_index) > 0) {
    text <- paste(paragraphs[(start_index[1] + 1):length(paragraphs)], collapse = " ")
  } else {
    text <- ""
    warning("Start keyword not found in the document.")
  }
  return(text)
}

# Function to count keyword occurrences in the text
count_keywords <- function(text, keywords) {
  keyword_counts <- sapply(keywords, function(keyword) {
    pattern <- paste0("\\b", keyword ,"\\b")
    #str_count(text, fixed(keyword, ignore_case = TRUE))
    str_count(text, regex(pattern, ignore_case = TRUE))
  })
  return(keyword_counts)
}

# Function to plot histogram of keyword counts
plot_histogram <- function(keyword_counts, title_text, type) {
  # Create a data frame for plotting
  df <- data.frame(
    Keyword = names(keyword_counts),
    Frequency = as.numeric(keyword_counts)
  )
  totals = sum(keyword_counts)#[which(df$Frequency > 0)]

  # Plot the histogram
  ggplot(df[which(df$Frequency>3),], aes(x = Keyword, y = Frequency)) +
    geom_bar(stat = "identity", fill = "skyblue", color = "black") +
    geom_text(aes(label = Frequency), vjust = -0.5) +
    labs(title = title_text, x = "Keywords", y = "Frequency") +
    theme_minimal() + 
    theme(plot.title = element_text(hjust = 0.5),
          axis.text.x = element_text(angle=45, hjust =1)) +
     geom_text(aes(x = Inf, y = Inf, label = paste("Sum:", round(totals, 2))),
              hjust = 1,
              vjust = 1,
              position = "dodge",
              color = "black",
              size = 3) +
    geom_text(aes(x=1, y=Inf, label = paste("Percent:", round(type, 2), "%")),
              hjust = "inward",
              vjust = "inward",
              position = "dodge",
              color = "black",
              size = 3.5)
}

#Find percent of Communal and Agentic keywords
keyword_percentage <- function(keyword_counts_Agentic, keyword_counts_Communal){
  df_Agentic <- data.frame(
    Keyword = names(keyword_counts_Agentic),
    Frequency = as.integer(keyword_counts_Agentic),
    Type = c("Agentic")
  )
  
  df_Communal <- data.frame(
    Keyword = names(keyword_counts_Communal),
    Frequency = as.integer(keyword_counts_Communal),
    Type = c("Communal")
  )
  total <- rbind(df_Agentic, df_Communal)
  
  total1 <- filter(total, Frequency > 0)
  
  percent <- setDT(total1)[ , 100 * .N / nrow(total1), by = Type ]
  
  total2 <- total1 %>%
    group_by(Type) %>%
    summarise(Frequency = sum(Frequency)) %>%
    mutate(Percent = prop.table(Frequency) * 100) %>%
    ungroup
  
  # x = 1:16
  # ggplot(total2, aes(x = Type, y = Frequency))+
  #          geom_bar(stat = "identity", fill = "skyblue", color = "black")
  # 
  #total2
  }

#File path or name of the document in your directory that you want to scan use
file_path <- "Rachel_PD.docx"

# Keyword to start scanning from
start_keyword <- "Demonstrated originality" # Replace with the specific section heading or keyword
# Keywords to count ----
keywords <- c('I', 'my', 'ability', 'able', 'absolute', 'accelerate', 'accelerated', 'accomplish', 'accomplishment', 'achieve', 'achievement',
              'acquired', 'act', 'action', 'active', 'actualize', 'administer', 'advance', 'advanced',
              'adventurous', 'advise', 'aggressive', 'ambition', 'amplified', 'appraise', 'army', 'aspire',
              'assert', 'asserted', 'assertible', 'assess', 'assign', 'assured', 'attain', 'attained', 'authored',
              'authoritative', 'authority', 'authorize', 'authorized', 'autonomous', 'awarded', 'best',
              'better', 'blocked', 'boast', 'bold', 'boost', 'boosted', 'brave', 'brave', 'build', 'builder',
              'built', 'campaign', 'campaigned', 'capable', 'certain', 'certify', 'chaired', 'challenge', 'change', 'changed',
              
              'charge', 'charted', 'coach', 'coaching', 'command', 'commanded', 'compel', 'compelled', 'competent', 'competitive',
              'complete', 'completed', 'compose', 'composed', 'comprehend', 'conceptualize', 'conclude',
              'conduct', 'confidence', 'confident', 'conflict', 'conflicts', 'consult', 'contest',
              'control', 'controlled', 'controlling', 'convince', 'convinced', 'correct', 'courage', 'courageous',
              'create', 'created', 'credential', 'credible', 'criticize', 'critiqued', 'daring',
              'decide', 'deciding', 'decision', 'decision-making', 'decisive', 'decisiveness',
              'delegate', 'deliver', 'demand', 'demanded', 'deploy', 'design', 'designate', 'determination',
              'determined', 'direct', 'directing', 'direction', 'directive', 'directorial',
              'disciplined', 'discovered', 'dispute', 'do', 'dominance', 'dominant', 'dominate',
              'dominating', 'drove', 'earn', 'earned', 'educate', 'effective', 'efficient', 'effort',
              'enact', 'enforced', 'engineered', 'enlist', 'enterprise', 'enterprising', 'entrepreneur',
              'establish', 'established', 'evaluate', 'evaluation', 'exceeded', 'excel', 'execute',
              'executed', 'executive', 'expedite', 'expedited', 'expert', 'fearless', 'financial',
              'finish', 'focus', 'focused', 'force', 'forceful', 'forge', 'formalize', 'formed',
              'forward', 'founded', 'freedom', 'furthered', 'gain', 'generate', 'get information',
              'goal', 'goal-oriented', 'govern', 'handle', 'hardworking', 'headed', 'hired', 'implement',
              'important', 'impose', 'imposing', 'impressive', 'improve', 'improved', 'in charge',
              'independence', 'independent', 'individualistic', 'induce', 'industrious', 'influenced',
              'initiating', 'initiative', 'initiate', 'innovative', 'insight', 'insist', 'insistent',
              'inspect', 'inspire', 'instigate', 'instructed', 'instrumental', 'intellectual',
              'intelligent', 'inventiveness', 'investigate', 'investigative', 'judge', 'judged',
              'judgment', 'Knowledgeable', 'launched', 'lead', 'leader', 'leadership', 'leading',
              'lobbied', 'logic', 'logical', 'make decisions', 'manage', 'masculine', 'mastery',
              'maximize', 'maximize returns', 'maximized', 'methodical', 'meticulous', 'military',
              'mobilized', 'motivate', 'negotiated', 'obtain', 'officer', 'operate', 'operations',
              'oppose', 'optimize', 'orchestrated', 'order', 'outpaced', 'outperformed', 'outspoken',
              'overhauled', 'oversaw', 'oversee', 'own', 'perform', 'persevering', 'persistence',
              'persistent', 'persuaded', 'pioneer', 'pioneered', 'plan', 'prepare', 'presented',
              'presiding', 'pressure', 'prioritize', 'proactive', 'problem solve', 'procure', 'produce',
              'productive', 'production', 'professional', 'proficient', 'programming', 'progress', 'project',
              'promoted', 'proud', 'proven', 'published', 'purposefulness', 'pursued', 'qualified',
              'rational', 'rationality', 'realistic', 'recommended', 'recruit', 'recruited', 'redefine',
              'reengineered', 'refocused', 'regulate', 'regulated', 'represented', 'require', 'resolute',
              'resolve', 'resourceful', 'Respected', 'responsible', 'restrict', 'results-driven', 'reviewed',
              'revised', 'savvy', 'screened', 'scrutinized', 'secure', 'secured', 'self-assured',
              'self-confident', 'self-reliant', 'self-starter', 'self-sufficient', 'sell', 'serious',
              'showcased', 'skill', 'skills', 'solve', 'spearhead', 'spearheaded', 'specify', 'standardized',
              'steadfast', 'steady', 'straightforward', 'strategic', 'streamlined', 'strength', 'strengthen',
              'stress tolerance', 'strive', 'strong', 'strong personality', 'strong-willed', 'stubborn',
              'substantial', 'succeeded', 'success', 'successful', 'superseded', 'supervise', 'supervised',
              'supervising', 'supervisory', 'sure', 'surpassed', 'talent', 'talented', 'targeted',
              'technical', 'tenacious', 'terminate', 'thorough', 'threat', 'thrive', 'tough', 'tracked',
              'train', 'trained', 'transcended', 'transformed', 'uncompromising', 'unyielding', 'upgraded',
              'urge', 'vigorous', 'willed', 'win', 'won', 'risks', 'accept', 'accommodate','accommodated', 'adapt',
              'adaptable', 'adherence', 'advocated', 'affectionate', 'affiliate', 'affiliated', 'agree', 'agreed', 'aid', 'align',
              'alliance', 'amenable', 'amiable', 'amicable', 'appreciate', 'appreciated', 'approachable', 'arbitrated',
              'artistic', 'assist', 'assisted', 'babysit', 'benevolent', 'calm', 'calmed', 'calmly', 'care', 'careful', 'caring', 'cater',
              'cautious', 'charitable', 'charity', 'cheer', 'childlike', 'clarified', 'clerical',
              'closeness', 'co-authored', 'collaborate', 'collaborated', 'colleague', 'collective', 'collegial', 'comfort',
              'committed', 'communicate', 'communicated', 'community', 'compassion', 'compassionate', 'complain', 'complained', 'comply',
              'compromise', 'concern', 'concerned', 'conciliate', 'conform', 'congenial', 'connect',
              'consensus', 'consider', 'considerate', 'contact', 'contributed', 'cooperate', 'cooperation',
              'cooperative', 'coordinated', 'cordial', 'corresponded', 'corresponding', 'counsel',
              'courteous', 'cultivated', 'culture', 'delicate', 'devote', 'diplomatic', 'emotion',
              'emotional', 'empathetic', 'empathic', 'empathy', 'encourage', 'ethical', 'etiquette',
              'express', 'expressive', 'facilitate', 'facilitated', 'familiar', 'family', 'favorable',
              'fearful', 'feeling', 'feelings', 'felt', 'feminine', 'flatterable', 'flexible', 'follow',
              'foster', 'fostered', 'fostering', 'friendly', 'gather', 'gave', 'generous', 'generously', 'genial',
              'gentle', 'gentleness', 'give', 'glad', 'good-natured', 'gracious', 'graciousness',
              'grateful', 'guided', 'gullible', 'harmony', 'harmonized', 'heart', 'help', 'helpful', 'helping',
              'hesitant', 'honest', 'hospitable', 'hosted', 'human resources', 'humane', 'humanitarian',
              'humble', 'include', 'incorporated', 'interact', 'interpersonal', 'intuitive', 'intuitiveness',
              'join', 'kind', 'kindly', 'kindness', 'lenient', 'liaison', 'listen', 'listener', 'listening',
              'loyal', 'mediate', 'mention', 'mentor', 'mentored', 'mild', 'mindful', 'modest', 'moral', 'morale',
              'mutual', 'nanny', 'neat', 'neatness', 'needs', 'neighborly', 'nurture', 'nurturing', 'obey',
              'obligations', 'oblige', 'obliging', 'observant', 'open', 'participate', 'partner', 'patient',
              'peacemaker', 'perceive', 'personable', 'personal', 'placid', 'pleasant', 'pleasing', 'pleasure',
              'polite', 'provide', 'provide help', 'providing', 'quiet', 'receiving', 'receptive', 'receptiveness',
              'recognize', 'recognizing', 'rehabilitated', 'relate', 'relation', 'relationship', 'relationships',
              'respond', 'responsive', 'safeguarded', 'sensitive', 'sensitivity', 'serene', 'serve', 'served', 'service',
              'servicing', 'share', 'sharing', 'shy', 'sociable', 'social', 'society', 'soft', 'soft spoken', 'suggested',
              'support', 'supportive', 'sympathetic', 'talk', 'taught', 'team-player', 'tender', 'tenderly', 'tenderness',
              'tentative', 'thoughtful', 'thoughtfulness', 'tolerant', 'tolerate', 'touching', 'tranquil',
              'trustworthy', 'understand', 'understanding', 'unified', 'unify', 'unite', 'united', 'unselfish',
              'valued', 'volunteer', 'warm', 'warmhearted', 'warmheartedness', 'warmth', 'we', 'welcome', 'welcomed', 'well-mannered', 'yielding')

keywords_Agentic <- c('I', 'my','ability', 'able', 'absolute', 'accelerate', 'accelerated', 'accomplish', 'accomplishment', 'achieve', 'achievement', 
                      'acquired', 'act', 'action', 'active', 'actualize', 'administer', 'advance', 'advanced', 
                      'adventurous', 'advise', 'aggressive', 'ambition', 'amplified', 'appraise', 'army', 'aspire',
                      'assert', 'asserted', 'assertible', 'assess', 'assign', 'assured', 'attain', 'attained', 'authored', 
                      'authoritative', 'authority', 'authorize', 'authorized', 'autonomous', 'awarded', 'best', 
                      'better', 'blocked', 'boast', 'bold', 'boost', 'boosted', 'brave', 'brave', 'build', 'builder', 
                      'built', 'campaign', 'campaigned', 'capable', 'certain', 'certify', 'chaired', 'challenge', 'change',
                      'charge', 'charted', 'coach', 'coaching', 'command', 'commanded', 'compel', 'compelled', 'competent', 'competitive', 
                      'complete', 'completed', 'compose', 'composed', 'comprehend', 'conceptualize', 'conclude',
                      'conduct', 'confidence', 'confident', 'conflict', 'conflicts', 'consult', 'contest',
                      'control', 'controlled', 'controlling', 'convince', 'convinced', 'correct', 'courage', 'courageous', 
                      'create', 'created', 'credential', 'credible', 'criticize', 'critiqued', 'daring', 
                      'decide', 'deciding', 'decision', 'decision-making', 'decisive', 'decisiveness',
                      'delegate', 'deliver', 'demand', 'demanded', 'deploy', 'design', 'designate', 'determination',
                      'determined', 'direct', 'directing', 'direction', 'directive', 'directorial',
                      'disciplined', 'discovered', 'dispute', 'do', 'dominance', 'dominant', 'dominate',
                      'dominating', 'drove', 'earn', 'earned', 'educate', 'effective', 'efficient', 'effort',
                      'enact', 'enforced', 'engineered', 'enlist', 'enterprise', 'enterprising', 'entrepreneur', 
                      'establish', 'established', 'evaluate', 'evaluation', 'exceeded', 'excel', 'execute', 
                      'executed', 'executive', 'expedite', 'expedited', 'expert', 'fearless', 'financial',
                      'finish', 'focus', 'focused', 'force', 'forceful', 'forge', 'formalize', 'formed',
                      'forward', 'founded', 'freedom', 'furthered', 'gain', 'generate', 'get information',
                      'goal', 'goal-oriented', 'govern', 'handle', 'hardworking', 'headed', 'hired', 'implement', 
                      'important', 'impose', 'imposing', 'impressive', 'improve', 'improved', 'in charge',
                      'independence', 'independent', 'individualistic', 'induce', 'industrious', 'influenced',
                      'initiating', 'initiative', 'initiate', 'innovative', 'insight', 'insist', 'insistent',
                      'inspect', 'inspire', 'instigate', 'instructed', 'instrumental', 'intellectual', 
                      'intelligent', 'inventiveness', 'investigate', 'investigative', 'judge', 'judged',
                      'judgment', 'Knowledgeable', 'launched', 'lead', 'leader', 'leadership', 'leading', 
                      'lobbied', 'logic', 'logical', 'make decisions', 'manage', 'masculine', 'mastery', 
                      'maximize', 'maximize returns', 'maximized', 'methodical', 'meticulous', 'military',
                      'mobilized', 'motivate', 'negotiated', 'obtain', 'officer', 'operate', 'operations',
                      'oppose', 'optimize', 'orchestrated', 'order', 'outpaced', 'outperformed', 'outspoken',
                      'overhauled', 'oversaw', 'oversee', 'own', 'perform', 'persevering', 'persistence', 
                      'persistent', 'persuaded', 'pioneer', 'pioneered', 'plan', 'prepare', 'presented', 
                      'presiding', 'pressure', 'prioritize', 'proactive', 'problem solve', 'procure', 'produce', 
                      'productive', 'production', 'professional', 'proficient', 'programming', 'progress', 'project', 
                      'promoted', 'proud', 'proven', 'published', 'purposefulness', 'pursued', 'qualified',
                      'rational', 'rationality', 'realistic', 'recommended', 'recruit', 'recruited', 'redefine', 
                      'reengineered', 'refocused', 'regulate', 'regulated', 'represented', 'require', 'resolute',
                      'resolve', 'resourceful', 'Respected', 'responsible', 'restrict', 'results-driven', 'reviewed',
                      'revised', 'savvy', 'screened', 'scrutinized', 'secure', 'secured', 'self-assured',
                      'self-confident', 'self-reliant', 'self-starter', 'self-sufficient', 'sell', 'serious',
                      'showcased', 'skill', 'skills', 'solve', 'spearhead', 'spearheaded', 'specify', 'standardized',
                      'steadfast', 'steady', 'straightforward', 'strategic', 'streamlined', 'strength', 'strengthen', 
                      'stress tolerance', 'strive', 'strong', 'strong personality', 'strong-willed', 'stubborn',
                      'substantial', 'succeeded', 'success', 'successful', 'superseded', 'supervise', 'supervised',
                      'supervising', 'supervisory', 'sure', 'surpassed', 'talent', 'talented', 'targeted', 
                      'technical', 'tenacious', 'terminate', 'thorough', 'threat', 'thrive', 'tough', 'tracked',
                      'train', 'trained', 'transcended', 'transformed', 'uncompromising', 'unyielding', 'upgraded',
                      'urge', 'vigorous', 'willed', 'win', 'won', 'risks')

keywords_Communal <- c('we', 'our', 'accept', 'accommodate', 'accommodated', 'adapt', 
                       'adaptable', 'adherence', 'advocated', 'affectionate', 'affiliate', 'affiliated', 'agree', 'agreed', 'aid', 'align',
                       'alliance', 'amenable', 'amiable', 'amicable', 'appreciate', 'appreciated', 'approachable', 'arbitrated',
                       'artistic', 'assist', 'assisted', 'babysit', 'benevolent', 'calm', 'calmed', 'calmly', 'care', 'careful', 'caring', 'cater',
                       'cautious', 'charitable', 'charity', 'cheer', 'childlike', 'clarified', 'clerical', 
                       'closeness', 'co-authored', 'collaborate', 'collaborated', 'colleague', 'collective', 'collegial', 'comfort', 
                       'committed', 'communicate', 'communicated', 'community', 'compassion', 'compassionate', 'complain', 'complained', 'comply', 
                       'compromise', 'concern', 'concerned', 'conciliate', 'conform', 'congenial', 'connect', 
                       'consensus', 'consider', 'considerate', 'contact', 'contributed', 'cooperate', 'cooperation', 
                       'cooperative', 'coordinated', 'cordial', 'corresponded', 'corresponding', 'counsel', 
                       'courteous', 'cultivated', 'culture', 'delicate', 'devote', 'diplomatic', 'emotion', 
                       'emotional', 'empathetic', 'empathic', 'empathy', 'encourage', 'ethical', 'etiquette', 
                       'express', 'expressive', 'facilitate', 'facilitated', 'familiar', 'family', 'favorable', 
                       'fearful', 'feeling', 'feelings', 'felt', 'feminine', 'flatterable', 'flexible', 'follow',
                       'foster', 'fostered', 'fostering', 'friendly', 'gather', 'gave', 'generous', 'generously', 'genial',
                       'gentle', 'gentleness', 'give', 'glad', 'good-natured', 'gracious', 'graciousness',
                       'grateful', 'guided', 'gullible', 'harmony', 'harmonized', 'heart', 'help', 'helpful', 'helping',
                       'hesitant', 'honest', 'hospitable', 'hosted', 'human resources', 'humane', 'humanitarian', 
                       'humble', 'include', 'incorporated', 'interact', 'interpersonal', 'intuitive', 'intuitiveness', 
                       'join', 'kind', 'kindly', 'kindness', 'lenient', 'liaison', 'listen', 'listener', 'listening',
                       'loyal', 'mediate', 'mention', 'mentor', 'mentored', 'mild', 'mindful', 'modest', 'moral', 'morale',
                       'mutual', 'nanny', 'neat', 'neatness', 'needs', 'neighborly', 'nurture', 'nurturing', 'obey', 
                       'obligations', 'oblige', 'obliging', 'observant', 'open', 'participate', 'partner', 'patient', 
                       'peacemaker', 'perceive', 'personable', 'personal', 'placid', 'pleasant', 'pleasing', 'pleasure', 
                       'polite', 'provide', 'provide help', 'providing', 'quiet', 'receiving', 'receptive', 'receptiveness', 
                       'recognize', 'recognizing', 'rehabilitated', 'relate', 'relation', 'relationship', 'relationships', 
                       'respond', 'responsive', 'safeguarded', 'sensitive', 'sensitivity', 'serene', 'serve', 'served', 'service',
                       'servicing', 'share', 'sharing', 'shy', 'sociable', 'social', 'society', 'soft', 'soft spoken', 'suggested', 
                       'support', 'supportive', 'sympathetic', 'talk', 'taught', 'team-player', 'tender', 'tenderly', 'tenderness', 
                       'tentative', 'thoughtful', 'thoughtfulness', 'tolerant', 'tolerate', 'touching', 'tranquil',
                       'trustworthy', 'understand', 'understanding', 'unified', 'unify', 'unite', 'united', 'unselfish', 
                       'valued', 'volunteer', 'warm', 'warmhearted', 'warmheartedness', 'warmth', 'welcome', 'welcomed', 'well-mannered', 'yielding')

#Extract text from word document ----
#text <- tolower(text)
text <- extract_text_from_section(file_path, start_keyword)

#Count each keyword type
keyword_counts <- count_keywords(text, keywords)
keyword_counts_Agentic <- count_keywords(text, keywords_Agentic)
keyword_counts_Communal <- count_keywords(text, keywords_Communal)

#Get Keyword Percentages
keyword_percentage(keyword_counts_Agentic, keyword_counts_Communal)
percent_agentic = as.numeric(keyword_percentage(keyword_counts_Agentic, keyword_counts_Communal)[1,3])
percent_communal = as.numeric(keyword_percentage(keyword_counts_Agentic, keyword_counts_Communal)[2,3])

#Plot histograms
#plot_histogram(keyword_counts, "Total Keyword Frequency")
plot_histogram(keyword_counts_Agentic, "Agentic Keyword Frequency", percent_agentic)
plot_histogram(keyword_counts_Communal, "Communal Keyword Frequency", percent_communal)

#Top words ----
data_top5 <- df %>%
  arrange(desc(Frequency)) %>%
  head(5)

ggplot(data_top5, aes(x = Keyword, y = Frequency)) +
  geom_bar(stat="identity")+
  geom_text(aes(label = Frequency), vjust = -0.5) +
  geom_col() +
  labs(title = "", x = "Keyword", y = "Value") +
  theme_bw()
