
from optparse import OptionParser
import os, logging
import utils
import copy
import collections


# create model for mostlikely tag
def create_model(sentences):
    ## YOUR CODE GOES HERE: create a model
    model = collections.defaultdict(str)
    counts = collections.defaultdict(int)
    tagsperwordcoll = collections.defaultdict(str)
    allDistinctTags = []
    for sentence in sentences:
        for token in sentence:
            word = token.word
            tag = token.tag
            if not tag in allDistinctTags:
                allDistinctTags.append(tag)
            counts[word + '/' + tag] += 1
            tag += '_'
            tagsineachwordstr = tagsperwordcoll[word]
            if not tag in tagsineachwordstr:
                tagsperwordcoll[token.word] += tag

    for word1 in tagsperwordcoll.keys():
        temptaglist = tagsperwordcoll[word1].split('_')
        max = 0
        maxtag = ''
        for temptag in temptaglist:
            if counts[word1 + '/' + temptag] > max:
                max = counts[word1 + '/' + temptag]
                maxtag = temptag
        model[word1] = maxtag
    return model

def predict_tags(sentences, model):
    ## YOU CODE GOES HERE: use the model to predict tags for sentences

    for sentence in sentences:
        for token in sentence:

            token.tag = model[token.word]
            if model[token.word] is '':
                token.tag = "NN"
    return sentences
#template to change tag x to tag y when previous tag is z
def template1(train_sents, predict_sents):
    rules = collections.defaultdict(int)
    
    assert len(train_sents) == len(predict_sents), "Gold and system don't have the same number of sentence"
    for sent in range(len(train_sents)):  # for sentence in sentences:
        assert len(train_sents[sent]) == len(predict_sents[sent]), "Different number of token in sentence:\n%s" % train_sents[sent]
        initialToken = 1
        for train_tok, predict_tok in zip(train_sents[sent], predict_sents[sent]): # for token in sentence
	    #goes into the loop for first token in a sentence
            if (initialToken == 1):
                initialToken = 0
                previousTag = predict_tok.tag
                continue
	    #goes into the loop when the tag for train and predict words are mismatched
            if (train_tok.tag != predict_tok.tag):
                temprule = (previousTag, predict_tok.tag, train_tok.tag) # save the tags x,y,z in array
                rules[temprule]+=1 #count the times same combination occured
            previousTag = predict_tok.tag
    return rules

#filters the rules by adding only those rules which increases the accuray
def reduced_rule1(rules, train_sents, predict_sents,accuracy):
    reduced_rules = []

    for temprule in rules.keys():
        if(rules[temprule]>500):
            copypredict_sents = copy.deepcopy(predict_sents)
            singlerule = []
            singlerule.append(temprule)
            newPrediction = apply_rule1(copypredict_sents,singlerule)
            newaccuracy = utils.calc_accuracy(train_sents,newPrediction)

            if(newaccuracy > accuracy):
                reduced_rules.append(temprule)

    return reduced_rules
#apply all the reduced rules to the corpus
def apply_rule1(sentences, rules):
    for rule in rules:
        beforeTag = rule[0]
        fromTag = rule[1]
        toTag = rule[2]
        for sentence in sentences:
            initialToken = 1
            for token in sentence:
                if (initialToken == 1):
                        initialToken = 0
                        previousTag = token.tag
                        continue
                if (token.tag == fromTag):
                    if (beforeTag == previousTag):
                        token.tag = toTag
                previousTag = token.tag
    return sentences

#template to change from x to y when the previous tag is z and next tag is p 
def template2(training_corpus, predictions):
    rules = collections.defaultdict(int)
    assert len(training_corpus) == len(predictions), "Gold and system don't have the same number of sentence"
    for sentence in range(len(training_corpus)):  # for sentence in sentences:
        assert len(training_corpus[sentence]) == len(predictions[sentence]), "Different number of token in sentence:\n%s" % training_corpus[sent]
        initialToken = 0
	oldtoken = ''
	previousTag = ''

        for train_tok, prediction_tok in zip(training_corpus[sentence], predictions[sentence]): # for token in sentence
            if (initialToken < 2):
                initialToken += 1
		oldtoken = previousTag
                previousTag = prediction_tok.tag
		previoustrain = train_tok.tag
				
	    elif previousTag != previoustrain and prediction_tok.tag == train_tok.tag :
		prediction_tag = previousTag
		train_tag = previoustrain
                next_predict_tag = prediction_tok.tag
                next_train_tag = train_tok.tag
                firstToken = 2
		templist = (oldtoken, prediction_tag, train_tag, next_predict_tag)				#[z,x,y,p]
		#print templist
		rules[templist]+= 1
	    oldtoken = previousTag
	    previousTag=prediction_tok.tag
	    previostrain=train_tok.tag
				

    return rules

def reduced_rule2(rules, training_corpus, predictions,accuracy):
    reduced_rules = []

    for temprule in rules.keys():
        if(rules[temprule]>500):
            copypredictions = copy.deepcopy(predictions)
            singlerule = []
            singlerule.append(temprule)
            new_predictions = apply_rule2(copypredictions,singlerule)
            newaccuracy = utils.calc_accuracy(training_corpus,new_predictions)

            if(newaccuracy > accuracy):
                reduced_rules.append(temprule)
    return reduced_rules
	
def apply_rule2(predictions, rules):
    for rule in rules:
        beforeTag = rule[0]
        fromTag = rule[1]
        toTag = rule[2]
	nextTag = rule[3]
        for sentence3 in predictions:
            for previoustoken, currenttoken, futureToken in zip(sentence3[0::2],sentence3[1::2],sentence3[2::2]):
		if previoustoken.tag == beforeTag :
			if (currenttoken.tag == fromTag):
				if (futureToken.tag == nextTag):
					#print futureToken
					currenttoken.tag = toTag
    return predictions



# template to change x to y if two tags before current tag is z
def template3(training_corpus,predictions):                                                         
    rules = collections.defaultdict(int)
    assert len(training_corpus) == len(predictions), "Gold and system don't have the same number of sentence"
    for sentence in range(len(training_corpus)):  # for sentence in sentences:
        assert len(training_corpus[sentence]) == len(predictions[sentence]), "Different number of token in sentence:\n%s" % training_corpus[sent]
        initialToken = 0
	previousTag = ''
	for train_tok, prediction_tok in zip(training_corpus[sentence], predictions[sentence]): # for token in sentence
		if (initialToken == 0):
                	initialToken =1
                	previouspreviousTag = previousTag
                	previousTag = prediction_tok.tag
                	continue
        	if (train_tok.tag != prediction_tok.tag):
                	temprule = (previouspreviousTag, prediction_tok.tag, train_tok.tag) #[z,x,y] 
			rules[temprule]+= 1
        	previouspreviousTag = previousTag
        	previousTag = prediction_tok.tag
	return rules
def reduced_rule3(rules, train_sents, predict_sents,accuracy):
    reduced_rules = []

    for temprule in rules.keys():
        if(rules[temprule]>500):
            copypredict_sents = copy.deepcopy(predict_sents)
            singlerule = []
            singlerule.append(temprule)
            newPrediction = apply_rule3(copypredict_sents,singlerule)
            newaccuracy = utils.calc_accuracy(train_sents,newPrediction)

            if(newaccuracy > accuracy):
                reduced_rules.append(temprule)

    return reduced_rules



def apply_rule3(sentences, rules):
    for rule in rules:
        beforeTag = rule[0]
        fromTag = rule[1]
        toTag = rule[2]
        for sentence in sentences:
            initialToken = 0
            previousTag = ''
            for token in sentence:
                if (initialToken < 2):
                        initialToken += 1
                        previouspreviousTag = previousTag
                        previousTag = token.tag
                        continue
                if (token.tag == fromTag):
                    if (previouspreviousTag == beforeTag):
                        token.tag = toTag
                previouspreviousTag = previousTag
                previousTag = token.tag
    return sentences

if __name__ == "__main__":
    usage = "usage: %prog [options] GOLD TEST"
    parser = OptionParser(usage=usage)

    parser.add_option("-d", "--debug", action="store_true",
                      help="turn on debug mode")

    (options, args) = parser.parse_args()
    if len(args) != 2:
        parser.error("Please provide required arguments")

    if options.debug:
        logging.basicConfig(level=logging.DEBUG)
    else:
        logging.basicConfig(level=logging.CRITICAL)


    rules = []

    training_file = args[0]
    training_sents = utils.read_tokens(training_file)  
    test_file = args[1]
    test_sents = utils.read_tokens(test_file)          

    model = create_model(training_sents)


    sents = utils.read_tokens(training_file)
    predictions = predict_tags(sents, model)
    accuracy = utils.calc_accuracy(training_sents, predictions)
    print "Accuracy in training before rules applied [%s sentences]: %s" % (len(sents), accuracy)



    rules = template1(training_sents, predictions)

    reduced_rules = reduced_rule1(rules,training_sents,predictions,accuracy)
    
    new_predictions = apply_rule1(predictions,reduced_rules) 
    
    accuracy = utils.calc_accuracy(training_sents, new_predictions)
    
    print "Accuracy in training after rule1 applied [%s sentences]: %s" % (len(sents), accuracy)

    test_sents1 = utils.read_tokens(test_file)
    predictions = predict_tags(test_sents1, model)
    accuracy = utils.calc_accuracy(test_sents, predictions)
    print "Accuracy in testing before rules [%s sentences]: %s" % (len(test_sents1), accuracy)


    template1_prediction = apply_rule1(predictions,reduced_rules)
    accuracy = utils.calc_accuracy(test_sents, template1_prediction)
    print "Accuracy in testing after rule1 applied [%s sentences]: %s" % (len(test_sents), accuracy)  


    
    rules = template2(training_sents, new_predictions)
    reduced_rules = reduced_rule2(rules, training_sents, new_predictions,accuracy)
    new_prediction = apply_rule2(new_predictions,reduced_rules)
    accuracy = utils.calc_accuracy(training_sents, new_prediction)
    print "Accuracy in training after rules2 applied [%s sentences]: %s" % (len(sents), accuracy)



    template3_prediction = apply_rule2(template1_prediction,reduced_rules)
    accuracy = utils.calc_accuracy(test_sents, template3_prediction)
    print "Accuracy in testing after rule2 applied [%s sentences]: %s" % (len(test_sents), accuracy)   




    rules = template3(training_sents, new_prediction)
    reduced_rules = reduced_rule3(rules, training_sents, new_prediction,accuracy)
    new_prediction = apply_rule3(new_prediction,reduced_rules)    
    accuracy = utils.calc_accuracy(training_sents, new_prediction)
    print "Accuracy in training after rules3 applied [%s sentences]: %s" % (len(sents), accuracy)


    template2_prediction = apply_rule3(template3_prediction,reduced_rules)
    accuracy = utils.calc_accuracy(test_sents, template3_prediction)
    print "Accuracy in testing after rule3 applied [%s sentences]: %s" % (len(test_sents), accuracy) 
 


    




