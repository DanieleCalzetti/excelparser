import string
from typing import List

class read_document:

    def __init__(self, sheet):
        self.sheet = sheet

    def find_header(self, cov_header, like_header, row_header, impre_header, hashtag_header):
        columns = list(string.ascii_uppercase)
        self.column_cov, self.column_hashtag, self.column_like, self.column_hashtag= "", "", "", ""
        for letter in columns:
            header = self.sheet["".join([letter, str(row_header)])].value
            if header == cov_header:
                self.column_cov = letter
            elif header == like_header:
                self.column_like = letter
            elif header == impre_header:
                self.column_impre = letter
            elif header == hashtag_header:
                self.column_hashtag = letter
        return self.column_impre, self.column_cov, self.column_hashtag, self.column_like

    def check_row_used(self, first_post):
        self.row_in_use = []
        count = int(first_post)
        row = self.sheet["".join(["A", str(count)])].value
        while row != None: 
            self.row_in_use.append(str(count))
            count += 1
            row = self.sheet["".join(["A", str(count)])].value
        self.row_in_use

    def find_top_post(self):
        total_cov = {}
        total_like = {}
        total_impre = {}       
        for num in self.row_in_use:
            cov_value = self.sheet["".join([self.column_cov, str(num)])].value
            total_cov[num] = cov_value
            total_cov_filtered = dict(filter(lambda item: item[1] is not None, total_cov.items()))

            like_value = self.sheet["".join([self.column_like, str(num)])].value
            total_like[num] = like_value
            total_like_filtered = dict(filter(lambda item: item[1] is not None, total_like.items()))

            impre_value = self.sheet["".join([self.column_impre, str(num)])].value
            total_impre[num] = impre_value
            total_impre_filtered = dict(filter(lambda item: item[1] is not None, total_impre.items()))

        cov_top_three = sorted(total_cov_filtered.items(), key=lambda x: x[1], reverse=True)[0:3]
        self.cov_top_three_dict = dict((x, y) for x, y in cov_top_three)

        like_top_three = sorted(total_like_filtered.items(), key=lambda x: x[1], reverse=True)[0:3]
        self.like_top_three_dict = dict((x, y) for x, y in like_top_three)

        impre_top_three = sorted(total_impre_filtered.items(), key=lambda x: x[1], reverse=True)[0:3]
        self.impre_top_three_dict = dict((x, y) for x, y in impre_top_three)

        return self.cov_top_three_dict, self.like_top_three_dict, self.impre_top_three_dict

    def compare_hashtag_top_three(self):
        row_cov = list(self.cov_top_three_dict)
        first_post_cov = list(self.sheet["".join([self.column_hashtag, str(row_cov[0])])].value.split())
        second_post_cov = list(self.sheet["".join([self.column_hashtag, str(row_cov[1])])].value.split())
        third_post_cov = list(self.sheet["".join([self.column_hashtag, str(row_cov[2])])].value.split())
        self.common_hashtag_cov = []
        for hashtag in first_post_cov:
            if hashtag in second_post_cov and hashtag not in self.common_hashtag_cov:
                self.common_hashtag_cov.append(hashtag)
            if hashtag in third_post_cov and hashtag not in self.common_hashtag_cov:
                self.common_hashtag_cov.append(hashtag)
        for hashtag in second_post_cov:
            if hashtag in third_post_cov and hashtag not in self.common_hashtag_cov:
                self.common_hashtag_cov.append(hashtag)

        row_like = list(self.like_top_three_dict)
        first_post_like = list(self.sheet["".join([self.column_hashtag, str(row_like[0])])].value.split())
        second_post_like = list(self.sheet["".join([self.column_hashtag, str(row_like[1])])].value.split())
        third_post_like = list(self.sheet["".join([self.column_hashtag, str(row_like[2])])].value.split())
        self.common_hashtag_like = []
        for hashtag in first_post_like:
            if hashtag in second_post_like and hashtag not in self.common_hashtag_like:
                self.common_hashtag_like.append(hashtag)
            if hashtag in third_post_like and hashtag not in self.common_hashtag_like:
                self.common_hashtag_like.append(hashtag)
        for hashtag in second_post_like:
            if hashtag in third_post_like and hashtag not in self.common_hashtag_like:
                self.common_hashtag_like.append(hashtag)

        row_impre = list(self.impre_top_three_dict)
        first_post_impre = list(self.sheet["".join([self.column_hashtag, str(row_impre[0])])].value.split())
        second_post_impre = list(self.sheet["".join([self.column_hashtag, str(row_impre[1])])].value.split())
        third_post_impre = list(self.sheet["".join([self.column_hashtag, str(row_impre[2])])].value.split())
        self.common_hashtag_impre = []
        for hashtag in first_post_impre:
            if hashtag in second_post_impre and hashtag not in self.common_hashtag_impre:
                self.common_hashtag_impre.append(hashtag)
            if hashtag in third_post_impre and hashtag not in self.common_hashtag_impre:
                self.common_hashtag_impre.append(hashtag)
        for hashtag in second_post_impre:
            if hashtag in third_post_impre and hashtag not in self.common_hashtag_impre:
                self.common_hashtag_impre.append(hashtag)

        return self.common_hashtag_cov,  self.common_hashtag_like, self.common_hashtag_impre 
      
    def compare_hashtag(self):
        self.common_tags = []
        for hashtag in self.common_hashtag_cov:
            if hashtag in self.common_hashtag_like and hashtag not in self.common_tags:
                self.common_tags.append(hashtag)
            if hashtag in self.common_hashtag_impre and hashtag not in self.common_tags:
                self.common_tags.append(hashtag)
        for hashtag in self.common_hashtag_impre:
            if hashtag in self.common_hashtag_like and hashtag not in self.common_tags:
                self.common_tags.append(hashtag)
        self.str_common_tags = ' '.join([str(elem) for elem in self.common_tags])
        if self.str_common_tags == "":
            self.str_common_tags = "There aren't common hashtags"
        return self.str_common_tags
