# import difflib

# def string_similarity(self, a, b):
#     return difflib.SequenceMatcher(None, a, b).ratio()

# def fuzzy_similarity_score(self, target, candidate, string_threshold=0.85):
#     match_count = 0
#     total_score = 0

#     for t_item in target:
#         best_score = 0
#         for c_item in candidate:
#             score = self.string_similarity(t_item, c_item)
#             if score > best_score:
#                 best_score = score
#         if best_score >= string_threshold:
#             match_count += 1
#             total_score += best_score

#     similarity_ratio = match_count / len(target)
#     return similarity_ratio, total_score

# def find_most_similar_fuzzy(self, target, candidates, overall_threshold=0.7):
#     best_match = 0
#     best_score = 0
#     best_total_similarity = 0
#     current_index = 0

#     for idx, candidate in enumerate(candidates):
#         score, total_similarity = self.fuzzy_similarity_score(target, candidate)
#         #debugging statment to show the scores of each item in the candidates list
#         print(f"Candidate {idx}: score={score:.2f}, total_similarity={total_similarity:.2f}, candidate={candidate}")

#         if score == 1.0:
#             continue  # skip exact match

#         if score >= overall_threshold:
#             if score > best_score or (score == best_score and total_similarity > best_total_similarity):
#                 best_match = current_index
#                 best_score = score
#                 best_total_similarity = total_similarity
#         current_index +=1

#     print("\nMost similar (>=50% fuzzy match but <100%):", best_match)
#     return best_match


# target = ['SE AUTOT4_LS_51G', 'SE AUTOT4_N_51N', 'SE AUTOT4_LS_51P']

# nested_lists = [['PY400-50SE_ZG', 'PY400-50SE_ZP', 'test', 'PY400-50SE_51G', 'Posey-SE GOC'],
#  ['SE150SL_21G', 'SE150SL_21P', 'SE150SL_51Q'],
#  ['OL350-400SE_ZG', 'OL350-400SE_ZP'],
#  ['CH300-100SE_ZG', 'CH300-100SE_ZP'],
#  ['SE AUTOT4_HS_51P', 'SE AUTOT4_HS_51G'],
#  ['SE AUTOT3_HS_51G', 'SE AUTOT3_HS_51P'],
#  ['SE AUTOT4_LS_51G', 'SE AUTOT4_N_51N', 'SE AUTOT4_LS_51P'],
#  ['SE AUTOT3_LS_51P', 'SE AUTOT3_N_51N', 'SE AUTOT3_LS_51G'],
#  ['SE350HY_21P', 'SE350HY_21G'],
#  ['SE250-CH_21G', 'SE250-CH_21P', 'SE250-CH_51Q'],
#  ['SE450OL_21G', 'SE450OL_21P', 'SE450OL_51P', 'SE450OL_51Q'],
#  ['SE750POSEY_21G', 'SE750POSEY_21P'],
#  ['HY200-400SE_ZP', 'HY200-400SE_ZG'],
#  [],
#  ['SE050BT_ZG', 'SE050BT_ZP', 'SE TR HS 51P', 'SE050BT_51P', 'SE050BT_51Q'],
#  ['SE 15-M1 51P', 'SE 15-BT 51P', 'SE T1 LS 51N', 'SE 15-BT 51G'],
#  ['SL300-150SE_21P', 'SL300-150SE_21G']]

target = ['CO AUTOT3 51G HS', 'CO AUTO3 51P HS']

nested_lists = [['CO AUTOT3 51G HS', 'CO AUTO3 51P HS'], ['CO AUTOT3 51G LS', 'CO AUTOT3 51P LS']]

# Normalize target with 'HS' replacing 'LS'
def replace_ls_hs(self, item):
    if 'LS' in item:
        return item.replace('LS', 'HS')
    elif 'HS' in item:
        return item.replace('HS', 'LS')
    return item

def find_hs_or_ls(self, target, canadites):
    # Build the flipped version of LS/HS items
    flipped_items = [self.replace_ls_hs(i) for i in target if 'LS' in i or 'HS' in i]
    flipped_items = list(filter(None, flipped_items))  # Remove None

    # Now search for a list that contains all flipped items
    matching_list = []
    current_index = 0
    for lst in canadites:
        if any(item in lst for item in flipped_items):
            matching_list = lst
            break
        current_index += 1

    print("Matching list:", matching_list)
    return(current_index)


find_hs_or_ls()