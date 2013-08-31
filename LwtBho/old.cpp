//		//tokenize the body text into a list of strings, not which strings are internal to tags, <the text here>
//		wchar_t* bodyIndex = const_cast<wchar_t*>(wBody.c_str());
//		vector<wchar_t*> prts;
//		vector<int> isInternal;
//		
//		int numCharsUntil = -1;
////MessageBox(NULL, L"a", L"a", MB_OK);
//		do
//		{
//			numCharsUntil = wcscspn(bodyIndex, L"<");
//			if (numCharsUntil == 0)
//			{
//				isInternal.push_back(1);
//				numCharsUntil = wcscspn(bodyIndex, L">");
//			}
//			else
//			{
//				if (prts.size() <= 0)
//					isInternal.push_back(0);
//				else
//				{
//					wchar_t* found = wcsstr(prts.back(), L"<script");
//					wcsncmp(prts.back(), L"<script", 7) == 0  || wcsncmp(prts.back(), L"<style", 6) == 0 ? isInternal.push_back(1) : isInternal.push_back(0);
//				}
//
//				numCharsUntil = wcscspn(bodyIndex, L"<");
//				numCharsUntil--; //I add to catch the end tag above, so this let's that subsequent add be a no-op, relatively
//			}
//			
//			wchar_t* theString = new wchar_t[numCharsUntil + 2]; // +2, one for the null, one to catch the > char or to undo balancing subtraction
//			wcsncpy(theString, bodyIndex, numCharsUntil + 1);
//			theString[numCharsUntil + 1] = '\0';
//			prts.push_back(theString);
//			bodyIndex = &bodyIndex[numCharsUntil+1];
//		} while (bodyIndex[0] != '\0');
////MessageBox(NULL, L"b", L"b", MB_OK);
////mb(_T("374"));
//		//create vector of words from page in order
//
//
//
//
//		wstring theNewBody(L"<style>span.status0 {background-color: #ADDFFF} span.status1 {background-color: #F5B8A9} span.status2 {background-color: #F5CCA9} span.status3 {background-color: #F5E1A9} span.status4 {background-color: #F5F3A9} span.status5 {background-color: #DDFFDD} span.status99 {background-color: #F8F8F8;border-bottom: solid 2px #CCFFCC} span.status98 {background-color: #F8F8F8;border-bottom: dashed 1px #000000}</style>");
//		for (int i = 0; i < prts.size(); ++i) //only consider every other entry, for external-to-tag text
//		{
////mb("396");
//			//skip text that is internal to tags
//			if (isInternal[i] == 1)
//			{
//				//wordsInChunk.push_back(-1);
//				theNewBody.append(prts[i]);
//				continue;
//			}
//
//			//use reg to grab vector of words in chunk
//			wstring wprt(prts[i]);
//			wregex wrgx(WordChars, regex_constants::icase | regex_constants::optimize | regex_constants::ECMAScript);
//			vector<wstring> chunkWords;
//
//			regex_iterator<wstring::const_iterator, wchar_t,regex_traits<wchar_t>> regit(wprt.begin(), wprt.end(), wrgx);
//			regex_iterator<wstring::const_iterator, wchar_t,regex_traits<wchar_t>> rend;
//			while (regit != rend)
//			{
//				chunkWords.push_back(regit->str());
//				++regit;
//			}
//
//			if (chunkWords.size() <= 0) // no words present, pass on text unchanged
//			{
//				theNewBody.append(prts[i]);
////mb("422");
//				continue;
//			}
//
//
//			//get db values for all entries not yet in resultSet
////mb("423");
//
//			if (!tStmt.prepare(tConn, _T("select WoStatus from words where WoLgID = ? AND WoTextLC = ?")))
//			{
//				mb(_T("Cannot prepare query. 408wiehfs"));
//				mb(tConn.last_error());
//			}
//			
//			tStmt.param(1).set_as_string(tLgID);
//			for (int j = 0; j < chunkWords.size(); ++j)
//			{
//				wstring srcval(chunkWords[j]);
//				unordered_map<wstring,wstring>::iterator it = resultSet.find(wstring_tolower(srcval));
//				if (it != resultSet.end());
//				else
//				{
//					tStmt.param(2).set_as_string(srcval);
//					if(!tStmt.execute())
//					{
//						mb(_T("Could not execute statment. 429shyduhf"));
//						mb(tStmt.last_error());
//					}
//					else
//					{
//						resultSet.insert(unordered_map<wstring,wstring>::value_type(srcval, tStmt.fetch_next() ? tStmt.field(1).as_string() : L"0"));
//						tStmt.free_results();
//					}
//				}
//			}
//
//			//append to transformed body tag content
//			wchar_t* prtIndOne = prts[i];
//			wchar_t* prtIndTwo = prts[i];
//
//			for (int j = 0; j < chunkWords.size(); ++j)
//			{
//				prtIndTwo = wcsstr(prtIndOne, chunkWords[j].c_str());
//				if (prtIndOne != prtIndTwo)
//				{
//					wstring wnon(prtIndOne, 0, prtIndTwo - prtIndOne);
//					//wchar_t* nonWord = new wchar_t[prtIndTwo - prtIndOne + 1];
//					//wcsncpy(nonWord, prtIndOne, prtIndTwo - prtIndOne);
//					//nonWord[prtIndTwo - prtIndOne] = '\0';
//					//theNewBody.append(nonWord);
//					theNewBody.append(wnon);
//					//delete [] nonWord;
//				}
//				
//				theNewBody.append(L"<span class=\"status");
//				wstring srcval(chunkWords[j]);
//				transform(srcval.begin(), srcval.end(), srcval.begin(), towlower);
//				unordered_map<wstring,wstring>::iterator it = resultSet.find(srcval);
//				if (it == resultSet.end())
//					MessageBox(NULL, L"iterator at end unexpectedly", L"258ydttgcv", MB_OK);
//				theNewBody.append(it->second);
//				theNewBody.append(L"\">");
//				theNewBody.append(chunkWords[j]);
//				theNewBody.append(L"</span>");
//				prtIndOne = prtIndTwo + chunkWords[j].size();
//			}
//
//			// capture remained nonword fragment of chunk			
//			if (*prtIndOne != '\0')
//				theNewBody.append(prtIndOne);
//		}
//		isInternal.clear();
////MessageBox(NULL, L"c", L"c", MB_OK);
//		//free allocated wchar_t arrays stored in prts vector
//		for (int i = 0; i < prts.size(); ++i)
//			delete [] prts[i];
////MessageBox(NULL, L"d", L"d", MB_OK);


		//replace original HTML entities back into document
		//ReturnOriginalEntities(theNewBody, savedEntities);